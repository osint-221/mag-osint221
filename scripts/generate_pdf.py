import os
import sys
import subprocess
from pathlib import Path


def find_chrome():
    candidates = []
    env = os.environ
    pf = env.get('ProgramFiles')
    pf86 = env.get('ProgramFiles(x86)')
    la = env.get('LocalAppData')
    if pf:
        candidates.append(Path(pf) / 'Google' / 'Chrome' / 'Application' / 'chrome.exe')
    if pf86:
        candidates.append(Path(pf86) / 'Google' / 'Chrome' / 'Application' / 'chrome.exe')
    if la:
        candidates.append(Path(la) / 'Google' / 'Chrome' / 'Application' / 'chrome.exe')
    # Microsoft Edge as fallback
    if pf:
        candidates.append(Path(pf) / 'Microsoft' / 'Edge' / 'Application' / 'msedge.exe')
    if pf86:
        candidates.append(Path(pf86) / 'Microsoft' / 'Edge' / 'Application' / 'msedge.exe')

    # PATH lookup
    from shutil import which
    for exe in ('chrome.exe', 'msedge.exe', 'chrome', 'msedge'):
        p = which(exe)
        if p:
            candidates.append(Path(p))

    for p in candidates:
        try:
            if p and p.exists():
                return str(p)
        except Exception:
            continue
    return None


def chrome_print_to_pdf(chrome_path, html_path, out_path):
    # Ensure file:// URL with forward slashes
    file_url = Path(html_path).absolute().as_posix()
    if os.name == 'nt' and not file_url.startswith('/'):
        file_url = '/' + file_url
    file_url = 'file://' + file_url
    # Remove existing file if present to avoid lock conflicts
    try:
        if os.path.exists(out_path):
            os.remove(out_path)
    except Exception:
        # retry a few times
        import time
        for i in range(5):
            try:
                time.sleep(0.2)
                if os.path.exists(out_path):
                    os.remove(out_path)
                break
            except Exception:
                continue

    cmd = [chrome_path, '--headless=new', '--disable-gpu', f'--print-to-pdf={out_path}', file_url]
    print('Running:', ' '.join(cmd))
    subprocess.check_call(cmd)


def pyppeteer_pdf_via_server(html_path, out_path, port=0, format='A4', landscape=False, margin_mm=10, chrome_path=None):
    try:
        from pyppeteer import launch
    except Exception:
        print('pyppeteer not installed — installing (user scope)...')
        # try user install to avoid permission issues
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', '--user', 'pyppeteer'])
        from pyppeteer import launch

    # Serve the repository root via a temporary HTTP server so relative paths and
    # root-relative assets ("/...") load the same as in production.
    from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
    import socket
    import threading

    class _Handler(SimpleHTTPRequestHandler):
        # silence logging
        def log_message(self, format, *args):
            pass

    # find free port if port==0
    sock = socket.socket()
    sock.bind(('127.0.0.1', 0))
    free_port = sock.getsockname()[1]
    sock.close()

    server = ThreadingHTTPServer(('127.0.0.1', free_port), _Handler)
    server_thread = threading.Thread(target=server.serve_forever, daemon=True)
    cwd = os.getcwd()
    try:
        # serve from repo root
        repo_root = Path(__file__).resolve().parent.parent
        os.chdir(str(repo_root))
        server_thread.start()

        url = f'http://127.0.0.1:{free_port}/' + Path(html_path).name

        async def _render():
            # If local Chrome is provided, use it to avoid downloading Chromium
            launch_kwargs = {}
            if chrome_path:
                launch_kwargs['executablePath'] = chrome_path
                launch_kwargs['args'] = ['--no-sandbox', '--disable-gpu']
            browser = await launch(**launch_kwargs)
            page = await browser.newPage()
            # set viewport to a wide landscape-like size when requested
            if landscape:
                await page.setViewport({'width': 1400, 'height': 900})
            else:
                await page.setViewport({'width': 1200, 'height': 1600})
            # emulate screen so the page is rendered exactly as on the site
            try:
                await page.emulateMediaType('screen')
            except Exception:
                try:
                    await page.emulateMedia('screen')
                except Exception:
                    pass
            await page.goto(url, {'waitUntil': 'networkidle0', 'timeout': 60000})
            # Inject minimal CSS to force page breaks between `.page` sections
            # without changing on-screen styles.
            break_css = '''
                .magazine-container { overflow: visible !important; }
                .page { break-after: page; page-break-after: always; break-inside: avoid; }
                /* print-only pair container */
                .pdf-pair { display: block; }
                .pdf-pair .page { height: 50%; box-sizing: border-box; }
            '''
            try:
                await page.addStyleTag({'content': break_css})
            except Exception:
                await page.evaluate("(css) => { var s = document.createElement('style'); s.type='text/css'; s.appendChild(document.createTextNode(css)); document.head.appendChild(s); }", break_css)
            # small delay to allow fonts/images to settle
            import asyncio as _asyncio
            await _asyncio.sleep(0.5)
            # Rearrange DOM to put pages 5+6 and 7+8 onto single PDF pages.
            # Work in page context: indices are 0-based.
            try:
                await page.evaluate('''() => {
                    const container = document.querySelector('.magazine-container');
                    if (!container) return;
                    const secs = Array.from(container.querySelectorAll('section'));
                    const pairs = [[4,5],[6,7]]; // zero-based pairs: 5+6, 7+8
                    // iterate pairs in descending order to avoid index shifts
                    pairs.forEach(pair => {
                        const [i,j] = pair;
                        if (!secs[i] || !secs[j]) return;
                        const first = secs[i];
                        const second = secs[j];
                        const wrapper = document.createElement('div');
                        wrapper.className = 'pdf-pair';
                        // style wrapper to occupy one printed page
                        wrapper.style.pageBreakAfter = 'always';
                        // clone nodes so we don't lose styles
                        const a = first.cloneNode(true);
                        const b = second.cloneNode(true);
                        a.style.height = '50%'; a.style.overflow = 'hidden';
                        b.style.height = '50%'; b.style.overflow = 'hidden';
                        wrapper.appendChild(a);
                        wrapper.appendChild(b);
                        // insert wrapper before the first original
                        first.parentNode.insertBefore(wrapper, first);
                        // remove originals
                        first.parentNode.removeChild(first);
                        // note: after removing first, second index shifts, find and remove second if exists
                        const secsNow = Array.from(container.querySelectorAll('section'));
                        const secondNode = secsNow.find(s => s.innerHTML === second.innerHTML);
                        if (secondNode) secondNode.parentNode.removeChild(secondNode);
                    });
                }''')
            except Exception:
                pass

            # Move the community contact block and footer into a dedicated page (page 7)
            try:
                await page.evaluate('''() => {
                    const container = document.querySelector('.magazine-container');
                    const community = document.querySelector('.community-contact');
                    const footer = document.querySelector('.footer-minimal');
                    if (!container || !community) return;
                    const newSection = document.createElement('section');
                    newSection.className = 'page';
                    newSection.style.display = 'flex';
                    newSection.style.flexDirection = 'column';
                    newSection.style.justifyContent = 'center';
                    newSection.style.alignItems = 'center';
                    newSection.style.padding = '60px 50px';
                    newSection.style.minHeight = '100vh';

                    const pc = document.createElement('div');
                    pc.className = 'page-content';
                    pc.style.width = '100%';
                    pc.style.display = 'flex';
                    pc.style.flexDirection = 'column';
                    pc.style.alignItems = 'center';
                    pc.style.justifyContent = 'center';
                    pc.style.flex = '1 1 auto';

                    const communityClone = community.cloneNode(true);
                    communityClone.style.margin = '0 auto';
                    communityClone.style.display = 'flex';
                    communityClone.style.flexDirection = 'column';
                    communityClone.style.alignItems = 'center';
                    communityClone.style.justifyContent = 'center';
                    communityClone.style.width = '80%';

                    pc.appendChild(communityClone);

                    // footer clone positioned at bottom
                    if (footer) {
                        const footerClone = footer.cloneNode(true);
                        footerClone.style.position = 'absolute';
                        footerClone.style.left = '0';
                        footerClone.style.right = '0';
                        footerClone.style.bottom = '10px';
                        footerClone.style.textAlign = 'center';
                        footerClone.style.width = '100%';
                        newSection.appendChild(pc);
                        newSection.appendChild(footerClone);
                        newSection.style.position = 'relative';
                    } else {
                        newSection.appendChild(pc);
                    }

                    container.appendChild(newSection);
                    // remove originals
                    try { community.parentNode.removeChild(community); } catch(e) {}
                    try { if (footer) footer.parentNode.removeChild(footer); } catch(e) {}
                }''')
            except Exception:
                pass
            # Use explicit page dimensions to control landscape orientation
            if landscape:
                width = '297mm'
                height = '210mm'
            else:
                width = '210mm'
                height = '297mm'
            pdf_options = {
                'path': str(out_path),
                'printBackground': True,
                'width': width,
                'height': height,
                'margin': {
                    'top': f'{margin_mm}mm',
                    'bottom': f'{margin_mm}mm',
                    'left': f'{margin_mm}mm',
                    'right': f'{margin_mm}mm',
                }
            }
            await page.pdf(pdf_options)
            await browser.close()

        import asyncio
        asyncio.get_event_loop().run_until_complete(_render())

    finally:
        try:
            server.shutdown()
            server.server_close()
        except Exception:
            pass
        os.chdir(cwd)



def main():
    root = Path(__file__).resolve().parent.parent
    html = root / 'magazine.html'
    out = root / 'magazine.pdf'
    if not html.exists():
        print('magazine.html not found in repository root:', html)
        sys.exit(1)
    # Prefer pyppeteer with a local HTTP server to ensure assets and root-relative
    # paths load exactly like in a browser. This reproduces the on-screen layout.
    import time
    tmp_out = out.with_name(f'magazine_temp_{int(time.time())}.pdf')
    try:
        # remove any leftover tmp file
        try:
            if tmp_out.exists():
                tmp_out.unlink()
        except Exception:
            pass
        chrome_for_pyppeteer = find_chrome()
        # user requested horizontal (landscape) pages; remove margins and group pages
        pyppeteer_pdf_via_server(str(html), str(tmp_out), format='A4', landscape=True, margin_mm=0, chrome_path=chrome_for_pyppeteer)
        # move temp to final (overwrite if needed)
        try:
            if out.exists():
                out.unlink()
        except Exception:
            pass
        tmp_out.replace(out)
        print('PDF generated at', out)
        return
    except Exception as e:
        print('pyppeteer (server) failed:', e)

    # Fallback: try local Chrome CLI print-to-pdf
    chrome = find_chrome()
    if chrome:
        print('Found Chrome at', chrome)
        try:
            chrome_print_to_pdf(chrome, str(html), str(out))
            print('PDF generated at', out)
            return
        except subprocess.CalledProcessError as e:
            print('Chrome failed to generate PDF:', e)

    print('All methods failed to generate PDF.')
    sys.exit(2)


if __name__ == '__main__':
    main()
