# MeteoSaver v1.0

Here we present MeteoSaver v1.0, a machine-learning based software for the transcription of historical weather data.

## Directory structure

Below is the structure for this project.

```
├── README.md                                               <- This includes general information on the project
|                                                              and introduction of project structure and modules
|
├── OCR_HTR_models                                          <- Trained and serialized models, model predictions, or model summaries.
|                                                         
├── data
│   ├── 00_post1960_DRC_hydroclimate_datasheet_images       <- Sample images of climate data sheets from the INERA
|                                                              Yangambi archives, DRC, arranged in folders (representing station numbers)
|   ├── 01_metadata_INERA_stations                          <- Metadata for the INERA meteorological stations
│
├── docs                                                    <- Flow charts and keys for MeteoSaver included in Muheki et al. 2025.
|
├── results                                                 <- All results obtained from exceution of MeteoSaver v1.0 on the sample ten datasheets
|   |                                                          included in this repository as a Minimal Working Example                                                                                         
│   ├──01_pre_QA_QC_transcribed_hydroclimate_data           <- Original automatic transcription using MeteoSaver before QA/QC checks
|   |                                       
│   ├── 02_post_QA_QC_transcribed_hydroclimate_data         <- Transcription after using MeteoSaver QA/QC checks
|   |                                      
│   ├── 03_validation_transcibed_data                       <- Validation of the transcription using MeteoSaver (in comparison to manual transcription)
│   │
│   ├── 04_final_refined_daily_hydroclimate_data            <- Final refined (formatted) trancribed data, in .xlsx and .tsv (SEF) ready for upload
|   │
│   ├── 05_transient_transcription_output                   <- Transient outputs (during processing)
|
│   ├── 06_manually_transcribed_hydroclimate_data           <- Manually transcribed data (For validation purposes)
|
|
|                                       
├── src                                                     <- Modules (2-6): Transcribing code/scripts for MeteoSaver v1.0
│   ├── main.py                                             <- Main script to run all the modules 1-6 of MetoSaver (scripts)
|   |                                                          i.e. in order (i) configuration, (iI) image-preprocessing module, (iii) table and cell
|   |                                                          detection model, (iv) transcription, (v) quality assessment and control,
|   |                                                          and (vi) data formatting and upload
│   │
│   ├── image_preprocessing_module.py                       <- Script to carry out image preprocessing of the original scans
|   |                                                          of climate data records
│   │
│   ├── table_and_cell_detection_model.py                   <- Script to detect the table and cells from the already
|   |                                                          pre-processed images
│   │
│   ├── transcription.py                                    <- Script to detect the text within the detected cells using
|   |                                                          an Optical Character Recognition (OCR) or Handwritten Text
|   |                                                          Recognition (HTR) model of your choice.
|   │
│   ├── quality_assessment_and_quality_control.py           <- Script to perform QA/QC checks on the original automatically transcribed data
|   |                                       
│   ├── validation.py                                       <- Script to generates a visual comparison of daily maximum, minimum,
|   |                                                          and average temperatures between manually transcribed data and
|   |                                                          QA/QC checked transcribed data for a specific station
│   └── data_formatting_and_upload.py                       <- Script to select the confirmed data (from the QA/QC) and convert it both an excel file 
|                                                              and to the Station Exchange Format, as well plot timeseries per station
|   
│
├── Dockerfile                                              <- Docker install routine for a virtual environment
|
├──LICENSE                                                  <- Licence
|
├── configuration.ini                                       <- Module 1: Configuration. User-defined settings to ensure smooth running of MeteoSaver
|
├── environment.yml                                         <- The requirements file for reproducing the analysis environment
|                                                              Generated with `conda env export > environment.yml`
|
├── job_script.sh                                           <- Job script for HPC infrastructure users to run the software
|
└── setup.py                                                <- Make this project pip installable with `pip install -e` 

```

## Setup

Two ways of creating reproducible environments are provided, the general Conda environment and an isolated Docker environment based on a Conda base image.

> [!WARNING]
> It is adviced to work in isolated Docker environments in order to ensure reproducibility, future online deployments, but first and foremost security of your computer system. Pip and to a lesser degree Conda and their python environments are a known malware vector. Although the framework we present vets the loaded library we can not assure the safety of all dependencies created downstream. The use of the local non-containerized setup is therefore not recommended.

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
from `pip` based [security issues](https://www.bleepingcomputer.com/news/security/pypi-suspends-new-user-registration-to-block-malware-campaign/). Once build locally no further downloads 
will be required.

```
# In the main project directory run
docker build -f Dockerfile -t transcribing_drc_data_environment .
```

To spin up a docker image using:

```
docker run -it -v /local_data:/docker_data_dir transcribing_drc_data_environment
```

## Modules
The figure below represents the modules in MeteoSaver v1.0
![Schematic representation of the modules in MeteoSaver v1.0](https://github.com/VUB-HYDR/MeteoSaver/blob/8ff79a3c003f157138824f32c91d5e41aa34ac75/docs/Schematic%20representation%20of%20the%20modules%20in%20MeteoSaver%20v1.0.png)


## How to run MeteoSaver v1.0 
1. After setting up the python environment using the [environment.yml](https://github.com/VUB-HYDR/MeteoSaver/blob/5b775a3047a38b86836bfdc24718ee2064756400/environment.yml) file available on this repository, input your user-sepcific settings in the configuration module

## Authors
Derrick Muheki

Koen Hufkens

Bas Vercruysse

Krishna Kumar Thirukokaranam Chandrasekar

Wim Thiery


## Acknowledgements

This template was inspired by the [python_proj_template](https://github.com/pepaaran/python_proj_template).
