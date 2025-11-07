import sys
import time
import os
import re
import json
import zipfile
import tempfile
import shutil
from pathlib import Path, PurePosixPath
from urllib.parse import unquote
import xml.etree.ElementTree as ET

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait

# ---------- Utilities ----------

def zjoin(base: str, rel: str) -> str:
    """Join ZIP/EPUB paths using POSIX rules (forward slashes)."""
    if not base:
        return str(PurePosixPath(rel))
    return str(PurePosixPath(base) / rel)

def wait_for_load(driver, timeout=30):
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )

def safe_read_text(p: Path):
    """Try a few encodings commonly seen in EPUB assets."""
    for enc in ("utf-8", "utf-16", "cp1252", "latin-1"):
        try:
            return p.read_text(encoding=enc, errors="ignore")
        except Exception:
            continue
    # Fallback to binary->utf-8 ignore
    return p.read_bytes().decode("utf-8", errors="ignore")

# ---------- Converter ----------

class EpubToPdfConverter:

    def __init__(self, file_path, output_name=None):
        self.file_path = file_path
        self.output_name = output_name or os.path.splitext(os.path.basename(file_path))[0]
        self.tmpdir = None  # set in convert_epub()

        chrome_options = Options()
        settings = {
            "recentDestinations": [{
                "id": "Save as PDF",
                "origin": "local",
                "account": "",
            }],
            "selectedDestinationId": "Save as PDF",
            "version": 2
        }
        prefs = {
            'printing.print_preview_sticky_settings.appState': json.dumps(settings),
            'savefile.default_directory': os.getcwd(),
            'download.default_directory': os.getcwd(),
        }
        chrome_options.add_experimental_option('prefs', prefs)
        chrome_options.add_argument('--kiosk-printing')
        # chrome_options.add_argument("--headless=new")  # optional

        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.minimize_window()
        self.driver.implicitly_wait(5)

    def convert_epub(self):
        print("Processing EPUB file...")

        if not os.path.exists(self.file_path):
            print(f"Error: File not found at {self.file_path}")
            self._cleanup_driver()
            sys.exit(1)

        if os.path.splitext(self.file_path)[1].lower() != '.epub':
            print("Error: This converter only supports .epub files.")
            self._cleanup_driver()
            sys.exit(1)

        try:
            with zipfile.ZipFile(self.file_path, 'r') as epub:
                container_xml = epub.read('META-INF/container.xml').decode('utf-8', errors='ignore')
                opf_path = self._get_opf_path(container_xml)
                if not opf_path:
                    print("Error: Could not locate OPF (content) file in container.xml")
                    self._cleanup_all()
                    sys.exit(1)

                opf_path = str(PurePosixPath(opf_path))  # normalize
                opf_dir  = str(PurePosixPath(opf_path).parent)
                opf_content = epub.read(opf_path).decode('utf-8', errors='ignore')

                # Extract all to temp dir so Chrome can load resources
                self.tmpdir = Path(tempfile.mkdtemp(prefix="epub_"))
                epub.extractall(self.tmpdir)

            print(f"Extracted EPUB to: {self.tmpdir}")
            combined_html = self._build_combined_html_from_fs(self.tmpdir, opf_content, opf_dir)

            if not combined_html.strip():
                print("Error: No content could be extracted from EPUB (spine/manifest and fallback empty).")
                self._cleanup_all()
                sys.exit(1)

            self._save_as_pdf(combined_html)

        except Exception as e:
            print(f"Error processing EPUB: {e}")
            self._cleanup_all()
            sys.exit(1)
        finally:
            self._cleanup_all()

    # ---------- XML Parsing Helpers ----------

    def _get_opf_path(self, container_xml: str) -> str:
        """
        Parse META-INF/container.xml robustly with namespaces.
        Prefer the first rootfile of type application/oebps-package+xml, else first rootfile.
        """
        try:
            # Parse and handle namespaces (container is usually in 'urn:oasis:names:tc:opendocument:xmlns:container')
            root = ET.fromstring(container_xml)
            # Any namespace: use wildcard matching
            rootfiles = root.findall(".//{*}rootfile")
            if not rootfiles:
                return None
            # prefer correct media-type if present
            for rf in rootfiles:
                media = rf.attrib.get("media-type", "")
                if media.endswith("oebps-package+xml"):
                    return rf.attrib.get("full-path")
            # fallback to first
            return rootfiles[0].attrib.get("full-path")
        except ET.ParseError:
            # Very odd EPUB; try regex fallback
            m = re.search(r'full-path="([^"]+\.opf)"', container_xml, flags=re.I)
            return m.group(1) if m else None

    def _parse_opf(self, opf_content: str):
        """
        Parse OPF with namespaces and return:
        - manifest: dict(id -> (href, media-type))
        - spine_ids: ordered list of idrefs
        - base_dir: derived by caller
        """
        # Try to detect namespaces present in <package>
        # Common OPF namespaces
        ns = {
            'opf': 'http://www.idpf.org/2007/opf',
            'dc':  'http://purl.org/dc/elements/1.1/',
        }
        try:
            root = ET.fromstring(opf_content)
        except ET.ParseError:
            # strip default xmlns and retry (last resort)
            cleaned = re.sub(r'\sxmlns="[^"]+"', '', opf_content, count=1)
            root = ET.fromstring(cleaned)
            ns = {}  # no default ns now

        # If the root is in a default ns, tag looks like {ns}package
        def tag(local):
            return f".//{{*}}{local}"

        manifest = {}
        for item in root.findall(tag("item")):
            _id = item.attrib.get("id")
            href = item.attrib.get("href")
            mtype = item.attrib.get("media-type", "")
            if _id and href:
                manifest[_id] = (href, mtype)

        spine_ids = []
        for itemref in root.findall(tag("itemref")):
            ref = itemref.attrib.get("idref")
            if ref:
                spine_ids.append(ref)

        return manifest, spine_ids

    # ---------- Content Build ----------

    def _build_combined_html_from_fs(self, extracted_dir: Path, opf_content: str, opf_base: str):
        print("Extracting EPUB content (spine order)...")

        manifest, spine_ids = self._parse_opf(opf_content)

        # 1) Build ordered hrefs from spine (preferred)
        ordered_hrefs = []
        if spine_ids:
            for _id in spine_ids:
                if _id in manifest:
                    href, mtype = manifest[_id]
                    ordered_hrefs.append(href)
        else:
            print("Spine empty, checking manifest...")

        # 2) If spine empty or produced 0 HTML files, fall back to all HTML-like manifest items
        html_like = []
        if not ordered_hrefs:
            for _id, (href, mtype) in manifest.items():
                if href and (href.lower().endswith((".xhtml", ".html", ".htm")) or "html" in mtype.lower()):
                    html_like.append(href)
            if html_like:
                # keep original order as listed in OPF by not sorting
                ordered_hrefs = html_like
                print(f"Manifest HTML-like count: {len(ordered_hrefs)}")

        # 3) If still nothing, last-resort: walk the extracted directory for .xhtml/.html/.htm
        if not ordered_hrefs:
            print("Spine empty, fallback manifest count: 0; walking filesystem for HTML...")
            candidates = []
            for p in extracted_dir.rglob("*"):
                if p.is_file() and p.suffix.lower() in (".xhtml", ".html", ".htm"):
                    # Make path relative to opf base if possible
                    rel = p.relative_to(extracted_dir).as_posix()
                    candidates.append(rel)
            # Heuristic order: natural-ish by path
            candidates.sort()
            ordered_hrefs = candidates
            print(f"Filesystem HTML count: {len(ordered_hrefs)}")

        if not ordered_hrefs:
            return ""

        all_content = []
        for href in ordered_hrefs:
            try:
                # Decode URL-encoded hrefs in OPF (e.g., spaces -> %20)
                href_dec = unquote(href)
                zip_rel = zjoin(opf_base, href_dec) if opf_base else href_dec
                fs_path = extracted_dir / Path(zip_rel.replace('/', os.sep))
                if not fs_path.exists():
                    # Try without joining base (some OPFs use absolute-like hrefs relative to root)
                    fs_path = extracted_dir / Path(href_dec.replace('/', os.sep))
                content = safe_read_text(fs_path)
                cleaned = self._clean_epub_html(content, base_path=str(Path(zip_rel).parent))
                all_content.append(cleaned)
                print(f"Processed: {zip_rel}")
            except Exception as e:
                print(f"Warning: Could not process {href}: {e}")

        return "\n<div style='page-break-before: always;'></div>\n".join(all_content)

    def _clean_epub_html(self, content: str, base_path: str):
        """Strip outer XML/HTML wrappers and rewrite resource links to local files."""
        import html
        content = re.sub(r'<\?xml[^>]*\?>', '', content, flags=re.I)
        content = re.sub(r'\sxmlns="[^"]*"', '', content, flags=re.I)

        # Prefer body content if present
        body_match = re.search(r'<body[^>]*>(.*?)</body>', content, flags=re.S | re.I)
        if body_match:
            content = body_match.group(1)

        # --- NEW image + CSS fix ---
        def fix_src(m):
            src = html.unescape(m.group(1))
            if src.startswith(('http://', 'https://', 'data:', '#')):
                return f'src="{src}"'
            # normalize relative path
            local_path = Path(self.tmpdir / Path(base_path) / Path(src)).resolve()
            if local_path.exists():
                return f'src="file:///{str(local_path).replace(os.sep, '/')}"'
            # fallback: maybe src was absolute in EPUB root
            alt_path = Path(self.tmpdir / Path(src)).resolve()
            if alt_path.exists():
                return f'src="file:///{str(alt_path).replace(os.sep, '/')}"'
            # not found – leave as-is for debugging
            print(f"Warning: missing image {src}")
            return f'src="{src}"'

        def fix_href_css(m):
            href = html.unescape(m.group(1))
            if href.startswith(('http://', 'https://', '#')):
                return f'href="{href}"'
            local_path = Path(self.tmpdir / Path(base_path) / Path(href)).resolve()
            if local_path.exists():
                return f'href="file:///{str(local_path).replace(os.sep, '/')}"'
            alt_path = Path(self.tmpdir / Path(href)).resolve()
            if alt_path.exists():
                return f'href="file:///{str(alt_path).replace(os.sep, '/')}"'
            print(f"Warning: missing CSS {href}")
            return f'href="{href}"'

        content = re.sub(r'src="([^"]+)"', fix_src, content, flags=re.I)
        content = re.sub(r'href="([^"]+\.css)"', fix_href_css, content, flags=re.I)

        return content

    # ---------- Print to PDF ----------

    def _save_as_pdf(self, inner_content: str):
        print("Converting to PDF...")

        html_template = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>{self.output_name}</title>
<style>
  @page {{ size: auto; margin: 20mm; }}
  body {{
    font-family: Arial, sans-serif;
    margin: 20px;
    line-height: 1.6;
    font-size: 12px;
  }}
  img {{
    max-width: 100%;
    height: auto;
    display: block;
    margin: 10px auto;
  }}
  div[style*="page-break-before"] {{ margin: 0; }}
</style>
</head>
<body>
<h1 style="page-break-after: always;">{self.output_name}</h1>
{inner_content}
</body>
</html>"""

        html_file = f"{self.output_name}_converted.html"
        with open(html_file, 'w', encoding='utf-8') as f:
            f.write(html_template)
        print(f"Temporary HTML saved as: {html_file}")

        try:
            html_path = Path(html_file).resolve()
            file_url = f"file:///{str(html_path).replace(os.sep, '/')}"
            self.driver.get(file_url)
            wait_for_load(self.driver, timeout=30)

            self.driver.execute_script('window.print();')
            time.sleep(4)

            print(f"PDF conversion attempted for: {self.output_name}")
            print("If you don't see the PDF, ensure Chrome can save to this folder.")
        except Exception as e:
            print(f"Error during PDF conversion: {e}")
        finally:
            try:
                os.remove(html_file)
                print(f"Cleaned up temporary file: {html_file}")
            except:
                pass

    # ---------- Cleanup ----------

    def _cleanup_driver(self):
        try:
            self.driver.quit()
        except:
            pass

    def _cleanup_all(self):
        self._cleanup_driver()
        if self.tmpdir and self.tmpdir.exists():
            try:
                shutil.rmtree(self.tmpdir, ignore_errors=True)
            except:
                pass

# ---------- CLI ----------

if __name__ == "__main__":
    print("EPUB to PDF Converter")
    print("=" * 30)
    print("Converts DRM-free EPUBs to PDF using Chrome (Save as PDF).")
    print("=" * 30)

    file_path = input("Please enter the path to your .epub file: ").strip().strip('"')
    if not file_path:
        print("Error: No file path provided.")
        sys.exit(1)

    custom_name = input("Enter custom output name (optional, press Enter to use file name): ").strip()
    converter = EpubToPdfConverter(file_path, custom_name if custom_name else None)
    converter.convert_epub()

    print("Done. Check this folder for the generated PDF.")
