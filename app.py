# First, install the required packages by running this command in your terminal:
# pip install flask python-pptx requests pillow opencv-python

from flask import Flask, request, send_file
from pptx import Presentation
from pptx.util import Inches
import os
import tempfile
import requests
from PIL import Image
import io
import cv2  # OpenCV for getting actual video dimensions

app = Flask(__name__)

@app.route('/')
def index():
    return "Hello, World! This is the PPT Video Injector API."

@app.route('/upload', methods=['POST'])
def upload_files():
    data = request.get_json()
    if not data or 'ppt' not in data or 'slides' not in data:
        return "Missing PPT URL or slides data", 400

    ppt_url = data['ppt']
    slides_info = data['slides']

    if not isinstance(slides_info, list) or not slides_info:
        return "Invalid slides data", 400

    # Create temporary directory
    with tempfile.TemporaryDirectory() as temp_dir:
        pptx_path = os.path.join(temp_dir, 'input.pptx')
        output_path = os.path.join(temp_dir, 'output.pptx')

        # Download PPTX from URL
        ppt_response = requests.get(ppt_url)
        if ppt_response.status_code != 200:
            return "Failed to download PPTX", 400
        with open(pptx_path, 'wb') as f:
            f.write(ppt_response.content)

        # Load the presentation
        prs = Presentation(pptx_path)
        num_slides = len(prs.slides)
        if num_slides == 0:
            return "Presentation has no slides", 400

        # Process each slide-video pair
        for idx, info in enumerate(slides_info):
            if not isinstance(info, dict) or 'number' not in info or 'videoLink' not in info:
                return f"Invalid slide info at index {idx}", 400

            slide_num = info['number']
            video_url = info['videoLink']

            if not isinstance(slide_num, int) or slide_num < 1 or slide_num > num_slides:
                return f"Invalid slide number {slide_num}", 400

            video_path = os.path.join(temp_dir, f'input_video_{slide_num}.mp4')
            thumbnail_path = os.path.join(temp_dir, f'thumbnail_{slide_num}.png')

            # Download video from URL
            video_response = requests.get(video_url)
            if video_response.status_code != 200:
                return f"Failed to download video for slide {slide_num}", 400
            with open(video_path, 'wb') as f:
                f.write(video_response.content)

            # Get actual video dimensions using OpenCV
            try:
                cap = cv2.VideoCapture(video_path)
                video_width_px = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
                video_height_px = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
                cap.release()
                
                if video_width_px == 0 or video_height_px == 0:
                    return f"Could not determine video dimensions for slide {slide_num}", 400
                
                print(f"Video dimensions: {video_width_px}x{video_height_px}")
                
                # Create thumbnail from first frame
                cap = cv2.VideoCapture(video_path)
                ret, frame = cap.read()
                if ret:
                    # Convert BGR to RGB
                    frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                    img = Image.fromarray(frame_rgb)
                    # Resize thumbnail to reasonable size while maintaining aspect ratio
                    img.thumbnail((320, 240), Image.Resampling.LANCZOS)
                    img.save(thumbnail_path)
                else:
                    # Fallback: create a simple colored rectangle
                    img = Image.new('RGB', (320, 180), color='blue')
                    img.save(thumbnail_path)
                cap.release()
                
            except Exception as e:
                return f"Failed to process video for slide {slide_num}: {str(e)}", 400

            # Convert pixel dimensions to inches (assuming 96 DPI)
            video_width_in = video_width_px / 96.0
            video_height_in = video_height_px / 96.0

            # Get slide dimensions in inches
            slide_width_in = prs.slide_width / 914400
            slide_height_in = prs.slide_height / 914400

            # Position 3: Calculate position (you can modify this as needed)
            # Position 3 could mean center, or specific coordinates - I'll place it at center
            left_in = slide_width_in - video_width_in
            top_in = slide_height_in - video_height_in

            # Ensure video doesn't exceed slide boundaries
            if video_width_in > slide_width_in or video_height_in > slide_height_in:
                # Scale down proportionally if video is larger than slide
                scale_factor = min(slide_width_in / video_width_in, slide_height_in / video_height_in) * 0.9  # 90% of slide max
                video_width_in *= scale_factor
                video_height_in *= scale_factor
                
                # Recalculate position after scaling
                left_in = (slide_width_in - video_width_in) / 2
                top_in = (slide_height_in - video_height_in) / 2

            # Ensure position is not negative
            left_in = max(0, left_in)
            top_in = max(0, top_in)

            print(f"Video size in inches: {video_width_in:.2f}x{video_height_in:.2f}")
            print(f"Position: left={left_in:.2f}, top={top_in:.2f}")

            # Add video to the specified slide
            slide = prs.slides[slide_num - 1]
            slide.shapes.add_movie(
                video_path,
                left=Inches(left_in),
                top=Inches(top_in),
                width=Inches(video_width_in),
                height=Inches(video_height_in),
                poster_frame_image=thumbnail_path,
                mime_type='video/mp4'
            )

        # Save the modified presentation
        prs.save(output_path)

        # Return the file
        return send_file(output_path, as_attachment=True, download_name='modified.pptx')


if __name__ == '__main__':
    app.run(debug=True)