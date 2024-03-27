# Data Rescue Congo DRC

Here we undertake data transcription of millions of daily precipitation and temperature records 
collected within the Congo Basin.

![](docs/data_rescue_flowchart.png)


## Directory structure
Below is the structure for this project.
```
├── README.md                            <- This includes general information on the project
|                                           and introduction of project structure and modules
├── LICENSE
|
├── data
│   ├── 10_sample_different_images       <- Sample images of climate data sheets from the INERA
|                                           Yangambi archives, DRC.
|
├── models                               <- Trained and serialized models, model predictions, or model summaries. 
|
├── notebook                             <- Jupyter notebook for exploration of the code only,
|                                           as well examples of outputs from the jupyter notebooks
|                                           such as Ms Excel with OCR results and clipped images.
│
├── src                                  <- Trancribing code/scripts for use in this project.
│   ├── main.py                          <- Script to run all the different modules (scripts) below
|   |                                       i.e. in order (1) image preprocessing module, (2) table
|   |                                       detection model, and (3) transcription model
│   │
│   ├── image_preprocessing.py           <- Script to carry out preprocessing of the original scans
|   |                                       of climate data records
│   │
│   ├── table_detection_model.py         <- Script to detect the table cells from the already
|   |                                       pre-processed images
│   │
│   ├── transcription_model.py           <- Script to detect the text within the detected cells using
|   |                                       an Optical Character Recognition (OCR) or Handwritten Text
|   |                                       Recognition (HTR) model of your choice.               
│   │
│   └── output                           <- Folder with outputs from the transcription
│
├── trials                               <- On-going trials/adaptations to the code.            
│   └── handwritten_recognition.py       <- On-going trails for better recognition of handwritten text
│
├── environment.yml                      <- The requirements file for reproducing the analysis environment.
|                                           Generated with `conda env export > environment.yml`
├── setup.py                             <- Make this project pip installable with `pip install -e`
|
└── Dockerfile                           <- Docker install routine for a virtual environment

```

## Setup

Two ways of creating reproducible environments are provided, the general Conda environment and an isolated Docker environment based on a Conda base image.

### Conda

To create an environment which is consistent use the environment file after installing Miniconda.

```bash
conda env create -f environment.yml
```

Activate the working environment using:

```bash
conda activate transcribing_drc_data_environment
```

### Docker

The dockerfile included provides a Conda environment ([see here for docker install instructions](https://docs.docker.com/engine/install/)).
You can build this docker image using the below command. This will download all required
python components and packages, while safeguarding (sandboxing) your system
from `pip` based security issues. Once build locally no further downloads 
will be required.

```
# In the main project directory run
docker build -f Dockerfile -t transcribing_drc_data_environment .
```

To spin up a docker image using:

```
docker run -it -v /local_data:/docker_data_dir transcribing_drc_data_environment
```

## Authors
Derrick Muheki

Koen Hufkens

Bas Vercruysse

Krishna Kumar Thirukokaranam Chandrasekar


This template was inspired by the [python_proj_template](https://github.com/pepaaran/python_proj_template).


