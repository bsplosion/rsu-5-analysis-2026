"""Download audio from a YouTube video via Playwright + direct URL extraction."""
import asyncio
import sys
import requests
from pathlib import Path
from playwright.async_api import async_playwright

AUDIO_DIR = Path(__file__).parent / "audio"
AUDIO_DIR.mkdir(exist_ok=True)


async def download_via_browser(video_id: str, dest: Path) -> dict | None:
    """Use Playwright to extract the stream URL and download it in-browser."""
    url = f"https://www.youtube.com/watch?v={video_id}"

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        ctx = await browser.new_context(accept_downloads=True)
        page = await ctx.new_page()

        print(f"  Loading video page ...")
        await page.goto(url, wait_until="domcontentloaded")
        await page.wait_for_timeout(4000)

        result = await page.evaluate("""() => {
            try {
                const pr = ytInitialPlayerResponse;
                const sr = pr.streamingData;
                const title = pr.videoDetails ? pr.videoDetails.title : '';
                const duration = pr.videoDetails ? pr.videoDetails.lengthSeconds : '0';

                const muxed = (sr.formats || []).find(f => f.url && f.itag === 18);
                if (muxed) {
                    return {url: muxed.url, itag: 18, mimeType: muxed.mimeType, title, duration};
                }
                const any_fmt = (sr.formats || []).find(f => f.url);
                if (any_fmt) {
                    return {url: any_fmt.url, itag: any_fmt.itag, mimeType: any_fmt.mimeType, title, duration};
                }
                return {error: 'No direct URL found', title, duration};
            } catch(e) { return {error: e.toString()}; }
        }""")

        if not result or "error" in result:
            await browser.close()
            return result

        print(f"  Title: {result['title']}")
        print(f"  Duration: {int(result['duration']) // 60}m")

        # Download the file within the browser context using fetch API
        print(f"  Downloading via browser fetch ...")
        download_js = """async (url) => {
            const response = await fetch(url);
            if (!response.ok) throw new Error('HTTP ' + response.status);
            const blob = await response.blob();
            const arrayBuffer = await blob.arrayBuffer();
            return Array.from(new Uint8Array(arrayBuffer));
        }"""

        try:
            data = await page.evaluate(download_js, result["url"])
            with open(dest, "wb") as f:
                f.write(bytes(data))
            result["downloaded"] = True
            result["size"] = len(data)
        except Exception as e:
            # If in-browser fetch fails, try navigating to the URL directly
            print(f"  Browser fetch failed ({e}), trying navigation download ...")
            dl_page = await ctx.new_page()
            resp = await dl_page.goto(result["url"])
            if resp and resp.ok:
                body = await resp.body()
                with open(dest, "wb") as f:
                    f.write(body)
                result["downloaded"] = True
                result["size"] = len(body)
            else:
                result["error"] = f"Download failed: HTTP {resp.status if resp else 'no response'}"

        await browser.close()
        return result


def download_file(url: str, dest: Path, cookies: dict = None):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
        "Referer": "https://www.youtube.com/",
        "Origin": "https://www.youtube.com",
    }
    print(f"  Downloading to {dest.name} ...")
    resp = requests.get(url, stream=True, headers=headers, cookies=cookies)
    resp.raise_for_status()
    total = int(resp.headers.get("content-length", 0))
    downloaded = 0
    with open(dest, "wb") as f:
        for chunk in resp.iter_content(chunk_size=1024 * 1024):
            f.write(chunk)
            downloaded += len(chunk)
            if total:
                pct = downloaded / total * 100
                print(f"\r  {downloaded / 1024 / 1024:.1f} / {total / 1024 / 1024:.1f} MB ({pct:.0f}%)", end="")
    print()


def extract_audio(mp4_path: Path, opus_path: Path, ffmpeg_dir: str = ""):
    import subprocess

    ffmpeg = "ffmpeg"
    if ffmpeg_dir:
        ffmpeg = str(Path(ffmpeg_dir) / "ffmpeg")

    cmd = [ffmpeg, "-i", str(mp4_path), "-vn", "-acodec", "libopus", "-b:a", "48k", str(opus_path), "-y"]
    print(f"  Extracting audio with ffmpeg ...")
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"  ffmpeg error: {result.stderr[:200]}")
        return False
    mp4_path.unlink()
    return True


async def main():
    video_id = sys.argv[1] if len(sys.argv) > 1 else "V3oYVFAx_6w"

    import glob
    ffmpeg_dirs = glob.glob(r"C:\Users\brand\AppData\Local\Microsoft\WinGet\Packages\Gyan.FFmpeg*\ffmpeg-*\bin")
    ffmpeg_dir = ffmpeg_dirs[0] if ffmpeg_dirs else ""

    mp4_path = AUDIO_DIR / f"{video_id}.mp4"
    opus_path = AUDIO_DIR / f"{video_id}.opus"

    print(f"Processing {video_id} ...")
    result = await download_via_browser(video_id, mp4_path)

    if not result or "error" in result:
        print(f"Failed: {result}")
        return

    if not result.get("downloaded"):
        print("Download did not complete")
        return

    print(f"  Downloaded: {result['size'] / 1024 / 1024:.1f} MB")

    if extract_audio(mp4_path, opus_path, ffmpeg_dir):
        print(f"  Audio: {opus_path.stat().st_size / 1024 / 1024:.1f} MB")
        print("Done!")
    else:
        print("  Audio extraction failed, keeping mp4")


if __name__ == "__main__":
    asyncio.run(main())
