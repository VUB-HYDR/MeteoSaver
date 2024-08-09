import os
import pandas as pd
import easyocr
from easyocr.utils import download_and_unzip

# Define the paths
image_dir = r'C:\Users\dmuheki\Downloads\CC-MNIST\MNIST + cobedore (Tesseract)'
label_dir = r'C:\Users\dmuheki\Downloads\CC-MNIST\MNIST + cobedore (Tesseract)'
csv_file = r'C:\Users\dmuheki\Downloads\dataset.csv'

# Prepare the data
data = []
for image_name in os.listdir(image_dir):
    if image_name.endswith(('.png', '.jpg', '.jpeg')):
        image_path = os.path.join(image_dir, image_name)
        # label_path = os.path.join(label_dir, image_name.replace('.png', '.gt'))
        base_name, ext = os.path.splitext(image_name)
        label_path = os.path.join(label_dir, base_name + '.gt.txt')
        
        if os.path.exists(label_path):
            with open(label_path, 'r', encoding='utf-8') as f:
                label = f.read().strip()
            data.append([image_path, label])

# Save to CSV
df = pd.DataFrame(data, columns=['image_path', 'label'])
df.to_csv(csv_file, index=False)



# Ensure the dataset CSV is correctly prepared
dataset_csv = csv_file
model_save_path = r'C:\Users\dmuheki\Downloads\easyocr_cobedore_model'

# Download necessary files for training English model
download_and_unzip('https://github.com/JaidedAI/EasyOCR/releases/download/v1.3/recog/latin.zip', 'latin.zip')

# Initialize the trainer for English
trainer = easyocr.Trainer(
    csv_file=dataset_csv,
    model_name='latin',
    save_model_dir=model_save_path,
    num_workers=4,
    batch_size=32,
    total_epoch=100
)

# Start training
trainer.train()


##### USE JUPYTER NOTEBOOK AND THIS https://github.com/JaidedAI/EasyOCR/blob/master/trainer/trainer.ipynb

