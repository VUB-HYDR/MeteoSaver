import os
from datetime import datetime

# Import all the required modules 
from transcribing_model.image_preprocessing_module import *
from transcribing_model.table_detection_model import *
from transcribing_model.transcription_model import *

# Setting up the current working directory; for both the input and output folders
cwd = os.getcwd()

# Directory for the original images (data sheets) to be transcribed
# ***Change this directory. Below we are testing this transcribing model on a folder with 10 sample images, here called 'data'
images_folder = os.path.join(cwd, 'data') #folder containing images
sample_images = os.path.join(images_folder, '10_sample_different_images') # 10 sample images

# TRIAL ON ONE TEST IMAGE FROM THE FOLDER
## This will be replaced with a 'for' loop after testing all the functions
one_test_image =  os.path.join(sample_images, '104_198104_SF_YAN.JPG')

# Module 1: Pre-processing the original images
preprocessed_image = image_preprocessing(one_test_image)

# Module 2: Table detection
detected_table_cells = table_detection(preprocessed_image)


# Module 3: Transcription / Handwritten Text Recognition
start_time = datetime.now() # Start recording transcribing time
ocr_model = 'Tesseract-OCR' # Selected OCR model out of: Tesseract-OCR, EasyOCR, PaddleOCR
transcribed_table = transcription(detected_table_cells, ocr_model)
end_time=datetime.now() # print total runtime of the code
print('Duration of transcribing: {}'.format(end_time - start_time))

