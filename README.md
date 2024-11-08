# MeteoSaver v1.0

Here we present MeteoSaver v1.0, a machine-learning based software for the transcription of historical weather data.

## Note: This README is still under development. However the code/scripts are up-to-date

![](docs/data_rescue_flowchart.png)


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
|                                                              included in this repository as a Minimal Working Example                            |                                                             
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
└── setup.py                                               <- Make this project pip installable with `pip install -e` 

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

## Pre-processing template matching

### Data sorting

Most historical tabulated data has a fixed format. This is a feature which I'll leverage later on. However, to ensure that the below procedures work well it is necessary to identify all the different table formats in a dataset. Particular care should be taken to ensure that small differences are accounted for, as even font changes can lead to less desirable post-processing results.

Overall make sure to:

- check for font differences
- check for line spacing differences
- overall correspondence between different tables should be high

In the Jungle Weather project we identified 20+ format of which three make up the bulk of all scanned data (>60%).

It is best to sort the images using a non-destructive method. It is therefore best to use a non-destructive photo editor or manager combined with tags rather to sort the data, rather than copying the source files around. In case of the Jungle Weather we used the [Shotwell photo editor and manager](https://wiki.gnome.org/Apps/Shotwell), on Windows and OSX Adobe Lightroom might serve the same purpose.

### Template generation

Once all different table formats are identified empty `templates` should be generated, and matching table cells annotated.

#### Creating an empty template

Where possible search the dataset for an already empty table. If no empty table exist use a table with as few data points as possible. Open this (almost) empty table using an image processing software. I suggest using [GIMP](https://www.gimp.org/), as I'll use a plugin later on to outline the cells of a table and it is freely available cross platforms.

Convert the this open file to a black and white template, while using the [levels](https://docs.gimp.org/2.10/en/gimp-tool-levels.html) and [curves]() to boost contrast and remove any unwanted gradients in the image. Remove all text which is not part of an empty template using the [eraser](https://docs.gimp.org/2.10/en/gimp-tool-eraser.html). The final result should look as the image below.

![](http://cobecore.org/images/documentation/mask.jpg)

When saving these templates use a comprehensive naming scheme with a prefix and a number separated with an underscore (_) such as: "format_1.jpg" corresponding to the folder containing the image data.

```
This formatting is important for successful use of the python processing code!
```

#### Outlining table cells

To specify the location of data within a table we will use the guides in GIMP, and a plugin to save this information. To save the guides in GIMP first install the ["save & load guides plugin"](https://github.com/khufkens/GIMP_save_load_guides). After installation of the plugin (and restarting GIMP) outline all cells in a table using GIMP guides. Below you see a template with all columns outlined with vertical guides.

![](http://cobecore.org/images/documentation/vertical_guides.png)

Once done, save the guides using the plugin (use: Image > Guides > Save). Make sure that the name used for the guides **exactly** matches the name of the image on which the guides are based. The guides will be saved in a file called "guides.txt" and stored this location:  "[userfolder]/.gimp-2.8/guides/guides.txt". Copy this file to your project folder for future processing (I store template data in a dedicated template folder containing all template images and the guides.txt file).

Note that you can save multiple sets of guides for multiple templates in the same guides.txt file.

## Post processing - Quality Assessement and Quality Control (QA/QC) checks:

![](https://github.com/VUB-HYDR/MeteoSaver/blob/b92a4cd9d6e1c41f818354e7abc7de434d624be9/docs/Post_processing_flowchart.png)


## Authors
Derrick Muheki

Koen Hufkens

Bas Vercruysse

Krishna Kumar Thirukokaranam Chandrasekar

Wim Thiery


## Acknowledgements

This template was inspired by the [python_proj_template](https://github.com/pepaaran/python_proj_template).
