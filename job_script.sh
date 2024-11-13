#!/bin/bash
#SBATCH --job-name=myjob
#SBATCH --time=100:00:00
#SBATCH --ntasks=1
#SBATCH --cpus-per-task=19
#SBATCH --partition=zen4
#SBATCH --mail-user=derrick.muheki@vub.be 
#SBATCH --mail-type=END,FAIL
#SBATCH --mem-per-cpu=10G
#SBATCH --output=slurm-%j.out
#SBATCH --error=slurm-%j.err

# set number of threads equal to 1
export OMP_NUM_THREADS=1

# set your hpc modules
module load Python/3.11.3-GCCcore-12.3.0 SciPy-bundle/2023.07-gfbf-2023a tesseract/5.3.4-GCCcore-12.3.0

# set the location of your pretrained language model for Tesseract OCR
export TESSDATA_PREFIX="/vscmnt/brussel_pixiu_data/_data_brussel/vo/000/bvo00012/vsc10520/OCR_HTR_models/"

# Activate virtual environment
source myenv/bin/activate

# Navigate to the directory containing your script
cd /vscmnt/brussel_pixiu_data/_data_brussel/vo/000/bvo00012/vsc10520/src

# Run the main Python script
python main.py