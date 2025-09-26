# First, install the required packages by running this command in your terminal:
# pip install flask python-pptx requests pillow

from flask import Flask, request, send_file
from pptx import Presentation
from pptx.util import Inches
import os
import tempfile
import requests
from PIL import Image
import io

app = Flask(__name__)

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

            video_path = os.path.join(temp_dir, f'input_video_{slide_num}.mp4')  # Assuming mp4, adjust if needed
            thumbnail_path = os.path.join(temp_dir, f'thumbnail_{slide_num}.png')

            # Download video from URL
            video_response = requests.get(video_url)
            if video_response.status_code != 200:
                return f"Failed to download video for slide {slide_num}", 400
            with open(video_path, 'wb') as f:
                f.write(video_response.content)

            # Extract video dimensions and create thumbnail using Pillow
            try:
                # For video dimensions, we'll use a default approach since Pillow doesn't handle video
                # We'll use standard HD dimensions (1920x1080) as fallback
                video_width_px = 1920
                video_height_px = 1080
                
                # Create a simple thumbnail using Pillow
                # Create a colored rectangle as a placeholder thumbnail
                img = Image.new('RGB', (320, 180), color='blue')
                img.save(thumbnail_path)
                
            except Exception as e:
                return f"Failed to process video for slide {slide_num}: {str(e)}", 400

            # Decide on video size in the slide (reduced max width from 6 to 4 inches for smaller size)
            max_slide_width_in = 2.0  # You can tweak this value to change the video width (in inches); smaller number = smaller video
            video_aspect = video_width_px / video_height_px if video_height_px != 0 else 1
            video_width_in = min(max_slide_width_in, video_width_px / 96)  # Assuming 96 DPI
            video_height_in = video_width_in / video_aspect  # Height adjusts automatically to maintain aspect ratio

            # Get slide dimensions in inches (standard widescreen is 13.333 x 7.5 inches)
            # But to be general, convert from EMU (914400 EMU per inch)
            slide_width_in = prs.slide_width / 914400
            slide_height_in = prs.slide_height / 914400

            # Calculate position: bottom-right aligned
            left_in = slide_width_in - video_width_in
            top_in = slide_height_in - video_height_in

            # Add video to the specified slide
            slide = prs.slides[slide_num - 1]
            slide.shapes.add_movie(
                video_path,
                left=Inches(left_in),
                top=Inches(top_in),
                width=Inches(video_width_in),
                height=Inches(video_height_in),
                poster_frame_image=thumbnail_path,
                mime_type='video/mp4'  # Adjust if your video is not mp4
            )

        # Save the modified presentation
        prs.save(output_path)

        # Return the file
        return send_file(output_path, as_attachment=True, download_name='modified.pptx')

if __name__ == '__main__':
    app.run(debug=True)