import os
from datetime import datetime

# Import all the required modules 
from transcribing_model.image_preprocessing_module import *
from transcribing_model.table_detection_model import *
from transcribing_model.transcription_model import *

## ***NEW
from transcribing_model.template_matching import *

# Setting up the current working directory; for both the input and output folders
cwd = os.getcwd()

# Directory for the original images (data sheets) to be transcribed
# ***Change this directory. Below we are testing this transcribing model on a folder with 10 sample images, here called 'data'
images_folder = os.path.join(cwd, 'data') #folder containing images and and template guides file
template_file_path = os.path.join(images_folder, 'guides_table.txt') # .txt file with the horizontal and vertical guides created for a template for the climate data sheets with GIMP*
sample_images = os.path.join(images_folder, '10_sample_different_images') # 10 sample images

# TRIAL ON ONE TEST IMAGE FROM THE FOLDER
## This will be replaced with a 'for' loop after testing all the functions
one_test_image =  os.path.join(sample_images, '701_19601_SF_YAN.JPG')
#one_test_image =  os.path.join(sample_images, '203_196503_SF_YAN.JPG')

# Module 1: Pre-processing the original images
preprocessed_image = image_preprocessing(one_test_image)

# Module 2: Table detection
detected_table_cells = table_detection(preprocessed_image)

# # Module 3: Transcription / Handwritten Text Recognition
# start_time = datetime.now() # Start recording transcribing time
# ocr_model = 'Tesseract-OCR' # Selected OCR model out of: Tesseract-OCR, EasyOCR, PaddleOCR
# transcribed_table = transcription(detected_table_cells, ocr_model)
# end_time=datetime.now() # print total runtime of the code
# print('Duration of transcribing: {}'.format(end_time - start_time))

## NEW. Template and boundary box creation module under testing. Module 3**: Template matching
start_time = datetime.now() # Start recording transcribing time
detected_table_image = detected_table_cells[2] #detected table from module 2
horizontal_guides, vertical_guides = parse_grid_file(template_file_path) # Template grid lines
template_grid_plot = plot_grid(detected_table_image, horizontal_guides, vertical_guides) ## Plot the grid lines on the detected table image. ** ONLY FOR VISUALIZATION PURPOSES
bounding_boxes = generate_bounding_boxes(horizontal_guides, vertical_guides)  # Here we generate bounding boxes (representing cells) for the detected table image
adjusted_bounding_boxes = adjust_bounding_boxes(bounding_boxes, height_increase= 10) # Here we increase height on both sides. We do this to ensure all text within the cell is captured especially for instances when the bounding box from the template cuts through a cell
plot_adjusted_bounding_boxes = plot_bounding_boxes(detected_table_image, adjusted_bounding_boxes) # Here, we plot the adjusted bounding boxes. ** ONLY FOR VISUALIZATION PURPOSES
end_time=datetime.now() # print total runtime of the code
print('Duration of transcribing: {}'.format(end_time - start_time))

## NEW.
