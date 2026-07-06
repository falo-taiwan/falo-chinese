import os
import sys
import argparse
import cv2
import requests
import numpy as np
from PIL import Image
from pptx import Presentation
from pptx.util import Inches
import yt_dlp

def format_timestamp(seconds: float) -> str:
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    s = int(seconds % 60)
    return f"{h:02d}:{m:02d}:{s:02d}"

def analyze_and_upload(video_url, task_id, cloudflare_url, threshold=8.0, interval=1.0):
    print(f"🎬 Starting YouTube extraction for URL: {video_url}")
    print(f"   Task ID: {task_id}")
    print(f"   Cloudflare URL: {cloudflare_url}")

    # 1. 下載 YouTube 影片 (優先下載 720p MP4)
    video_file = "temp_video.mp4"
    ydl_opts = {
        'format': 'bestvideo[height<=720][ext=mp4]+bestaudio[ext=m4a]/best[height<=720]/best',
        'outtmpl': video_file,
        'quiet': True,
        'no_warnings': True
    }
    
    title = "YouTube Video"
    duration = 0
    
    try:
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(video_url, download=True)
            title = info.get('title', 'YouTube Video')
            duration = info.get('duration', 0)
    except Exception as e:
        print(f"❌ Failed to download YouTube video: {e}")
        sys.exit(1)

    print(f"✅ Video downloaded successfully: '{title}' ({duration}s)")

    # 2. 向 Workers API 建立/初始化任務
    init_url = f"{cloudflare_url.rstrip('/')}/api/tasks"
    try:
        r = requests.post(init_url, json={
            "id": task_id,
            "title": title,
            "video_name": video_url,
            "video_duration": duration,
            "status": "analyzing"
        })
        print(f"📌 Task initialized in D1: {r.status_code}")
    except Exception as e:
        print(f"⚠️ Failed to init task in Cloudflare D1: {e}")

    # 3. 開啟影片進行 analysis
    cap = cv2.VideoCapture(video_file)
    if not cap.isOpened():
        print("❌ Cannot open video file with OpenCV.")
        sys.exit(1)

    fps = cap.get(cv2.CAP_PROP_FPS)
    total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
    print(f"📹 FPS: {fps}, Total Frames: {total_frames}")

    slide_candidates = []
    prev_low_res = None
    slide_no = 0

    # 建立暫存目錄存放投影片圖片
    temp_dir = "temp_slides"
    os.makedirs(temp_dir, exist_ok=True)

    # 以秒為間隔進行 MAE 掃描
    frame_step = int(fps * interval) if fps > 0 else 30
    
    for frame_idx in range(0, total_frames, frame_step):
        cap.set(cv2.CAP_PROP_POS_FRAMES, frame_idx)
        ret, frame = cap.read()
        if not ret:
            break
            
        current_time = frame_idx / fps if fps > 0 else 0
        
        # 轉成 32x18 灰階低解析度進行 MAE 計算
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        low_res = cv2.resize(gray, (32, 18))
        
        if prev_low_res is not None:
            # 計算 MAE 差異
            mae = np.mean(np.abs(low_res.astype(np.float32) - prev_low_res.astype(np.float32)))
            
            if mae > threshold:
                slide_no += 1
                timestamp = format_timestamp(current_time)
                print(f"📸 Detected new slide #{slide_no} at {timestamp} (MAE: {mae:.2f})")
                
                # 儲存高解析度 PNG 投影片
                slide_filename = os.path.join(temp_dir, f"slide_{slide_no:03d}.png")
                cv2.imwrite(slide_filename, frame)
                
                # 直傳此投影片至 Workers R2
                upload_url = f"{cloudflare_url.rstrip('/')}/api/tasks/{task_id}/upload"
                try:
                    with open(slide_filename, "rb") as f:
                        file_bytes = f.read()
                        
                    up_res = requests.post(upload_url, data=file_bytes, headers={
                        "X-Slide-No": str(slide_no),
                        "X-Timestamp": timestamp,
                        "X-Seconds": str(current_time)
                    })
                    print(f"   Upload slide #{slide_no} status: {up_res.status_code}")
                except Exception as e:
                    print(f"   ⚠️ Upload failed: {e}")
                    
                slide_candidates.append({
                    "slide_no": slide_no,
                    "timestamp": timestamp,
                    "seconds": current_time,
                    "path": slide_filename
                })
                
        else:
            # 第一張投影片
            slide_no += 1
            timestamp = format_timestamp(current_time)
            print(f"📸 Detected first slide #{slide_no} at {timestamp}")
            
            slide_filename = os.path.join(temp_dir, f"slide_{slide_no:03d}.png")
            cv2.imwrite(slide_filename, frame)
            
            # 直傳至 Workers
            upload_url = f"{cloudflare_url.rstrip('/')}/api/tasks/{task_id}/upload"
            try:
                with open(slide_filename, "rb") as f:
                    file_bytes = f.read()
                requests.post(upload_url, data=file_bytes, headers={
                    "X-Slide-No": str(slide_no),
                    "X-Timestamp": timestamp,
                    "X-Seconds": str(current_time)
                })
            except Exception as e:
                print(f"   ⚠️ Upload first slide failed: {e}")
                
            slide_candidates.append({
                "slide_no": slide_no,
                "timestamp": timestamp,
                "seconds": current_time,
                "path": slide_filename
            })
            
        prev_low_res = low_res

    cap.release()
    print(f"🎉 Analysis completed! Total slides captured: {len(slide_candidates)}")

    if not slide_candidates:
        print("⚠️ No slides detected.")
        sys.exit(0)

    # 4. 編譯 PDF 文件並上傳
    print("📄 Compiling PDF presentation...")
    pdf_path = "presentation.pdf"
    try:
        images_list = []
        for s in slide_candidates:
            img = Image.open(s["path"]).convert("RGB")
            images_list.append(img)
            
        if images_list:
            images_list[0].save(pdf_path, save_all=True, append_images=images_list[1:])
            print("   PDF compiled successfully.")
            
            # 上傳 PDF
            doc_url = f"{cloudflare_url.rstrip('/')}/api/tasks/{task_id}/document"
            with open(pdf_path, "rb") as f:
                requests.post(doc_url, data=f.read(), headers={
                    "X-Doc-Type": "pdf"
                })
            print("   PDF uploaded to Cloudflare R2.")
    except Exception as e:
        print(f"❌ Failed to compile/upload PDF: {e}")

    # 5. 編譯 PPTX 文件並上傳
    print("📊 Compiling PPTX presentation...")
    pptx_path = "presentation.pptx"
    try:
        prs = Presentation()
        # 設定 16:9 比例頁面尺寸
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        blank_slide_layout = prs.slide_layouts[6]
        
        for s in slide_candidates:
            slide = prs.slides.add_slide(blank_slide_layout)
            slide.shapes.add_picture(s["path"], 0, 0, width=prs.slide_width, height=prs.slide_height)
            
        prs.save(pptx_path)
        print("   PPTX compiled successfully.")
        
        # 上傳 PPTX
        with open(pptx_path, "rb") as f:
            requests.post(doc_url, data=f.read(), headers={
                "X-Doc-Type": "pptx"
            })
        print("   PPTX uploaded to Cloudflare R2.")
    except Exception as e:
        print(f"❌ Failed to compile/upload PPTX: {e}")

    # 6. 清理暫存檔案
    try:
        if os.path.exists(video_file):
            os.remove(video_file)
        for s in slide_candidates:
            if os.path.exists(s["path"]):
                os.remove(s["path"])
        if os.path.exists(temp_dir):
            os.rmdir(temp_dir)
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
        if os.path.exists(pptx_path):
            os.remove(pptx_path)
        print("🧹 Cleaned up all temporary files.")
    except Exception as e:
        print(f"⚠️ Error during cleanup: {e}")

    print("🏁 Execution finished successfully!")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="GitHub Actions YouTube Slide Extractor")
    parser.add_argument("--url", required=True, help="YouTube Video URL")
    parser.add_argument("--task-id", required=True, help="Cloudflare Task ID")
    parser.add_argument("--cloudflare-url", required=True, help="Cloudflare Worker Domain URL")
    parser.add_argument("--threshold", type=float, default=8.0, help="MAE change threshold")
    parser.add_argument("--interval", type=float, default=1.0, help="Frame check interval in seconds")
    
    args = parser.parse_args()
    analyze_and_upload(
        video_url=args.url,
        task_id=args.task_id,
        cloudflare_url=args.cloudflare_url,
        threshold=args.threshold,
        interval=args.interval
    )
