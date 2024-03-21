# Data_Rescue_Congo_DRC
Here we undertake data transcription of millions of daily precipitation and temperature records collected within the Congo Basin.

## Directory structure
Below is the structure for this project.
```
├── README.md                            <- This includes general information on the project and introduction of project structure and modules
├── LICENSE
|
├── data
│   ├── 10_sample_different_images       <- Sample images of climate data sheets from the INERA Yangambi archives, DRC.
│  
├── notebook                             <- Jupyter notebook for exploration of the code only, as well examples of outputs from the jupyter notebooks such as Ms Excel with OCR results and clipped images.
│
├── transcribing_model                   <- Trancribing code/scripts for use in this project.
│   ├── main.py                          <- Script to run all the different modules (scripts) below i.e. in order (1) image preprocessing module, (2) table detection model, and (3) transcription model
│   │
│   ├── image_preprocessing.py           <- Script to carry out preprocessing of the original scans of climate data records
│   │
│   ├── table_detection_model.py         <- Script to detect the table cells from the already pre-processed images
│   │
│   ├── transcription_model.py           <- Script to detect the text within the detected cells using an Optical Character Recognition (OCR) or Handwritten Text Tecognition (HTR) model of your choice.               
│   │
│   └── output                           <- Folder with outputs from the transcription
│
├── trails                               <- On-going trials/adaptations to the code.            
│   └── handwritten_recognition.py       <- On-going trails for better recognition of handwritten text
│
├── environment.yml                      <- The requirements file for reproducing the analysis environment. Generated with `conda env export > environment.yml`

```


## Python environment
To ensure reproducibility of our analysis, the [environment.yml](https://github.com/VUB-HYDR/Data_Rescue_Congo_DRC/blob/19af3b0897fc818428a8f503c2982c668b32eb54/environment.yml) provides a clone of our python environment, generated with `conda env export > environment.yml`, with all of its packages and versions. Users should use their terminal or an Anaconda Prompt to create their environment using this file.

## Authors
Derrick Muheki
Koen Hufkens
Bas Vercruysse
Krishna Kumar Thirukokaranam Chandrasekar


This template was inspired by the [python_proj_template](https://github.com/pepaaran/python_proj_template).


