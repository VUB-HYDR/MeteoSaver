
import os
import numpy as np
import cv2 as cv
import matplotlib.pyplot as plt


# # Setting up the current working directory; for both the input and output folders
# cwd = os.getcwd()

# # Directory for the original images (data sheets) to be transcribed
# # ***Change this directory. Below we are testing this transcribing model on a folder with 10 sample images, here called 'data'
# images_folder = os.path.join(cwd, 'data') #folder containing images and template guides file
# template_file_path = os.path.join(images_folder, 'guides_table.txt') # .txt file with the horizontal and vertical guides created for a template for the climate data sheets with GIMP*
# sample_images = os.path.join(images_folder, '10_sample_different_images') # 10 sample images


#PLOT GRID LINES
def parse_grid_file(file_path):
    with open(file_path, 'r') as file:
        lines = file.readlines()
        # Extracting the vertical and horizontal guide coordinates
        horizontal_guides_str = lines[0].split('|')[2]
        vertical_guides_str = lines[0].split('|')[3]
        # Split the guide coordinates by ','
        horizontal_guides = list(map(int, horizontal_guides_str.split(',')))
        vertical_guides = list(map(int, vertical_guides_str.split(',')))
    return horizontal_guides, vertical_guides

def plot_grid(image, horizontal_guides, vertical_guides):
    # Read the image
    #image = cv.imread(image_path)
    plt.imshow(image)
    # Plot horizontal lines
    for y in sorted(horizontal_guides):
        plt.axhline(y=y, color='r', linestyle='solid')
    # Plot vertical lines
    for x in sorted(vertical_guides):
        plt.axvline(x=x, color='r', linestyle='solid')
    plt.show()


def generate_bounding_boxes(horizontal_guides, vertical_guides):
    bounding_boxes = []
    for i in range(len(horizontal_guides) - 1):
        for j in range(len(vertical_guides) - 1):
            x = vertical_guides[j]
            y = horizontal_guides[i]
            width = vertical_guides[j + 1] - vertical_guides[j]
            height = horizontal_guides[i + 1] - horizontal_guides[i]
            # Ensure height is positive
            if height < 0:
                y += height  # Move the starting point up
                height = abs(height)  # Make height positive
            bounding_boxes.append((x, y, width, height))
    return bounding_boxes

#bounding_boxes = generate_bounding_boxes(horizontal_guides, vertical_guides)

def adjust_bounding_boxes(bounding_boxes, height_increase=10):
    adjusted_bounding_boxes = []
    for bbox in bounding_boxes:
        x, y, width, height = bbox
        # Increase height on both sides. We do this to ensure all text within the cell is captured especially for instances when the bounding box from the template cuts through a cell
        adjusted_y = y - height_increase
        adjusted_height = height + 2 * height_increase
        adjusted_bounding_boxes.append((x, adjusted_y, width, adjusted_height))
    return adjusted_bounding_boxes

#adjusted_bounding_boxes = adjust_bounding_boxes(bounding_boxes, height_increase=10)

# Plot only bounding boxes
def plot_bounding_boxes(image, bounding_boxes):
    #image = cv.imread(image_path)
    plt.imshow(image)
    ax = plt.gca()
    for bbox in bounding_boxes:
        x, y, width, height = bbox
        rect = plt.Rectangle((x, y), width, height, linewidth=0.1, edgecolor='r', facecolor='none')
        ax.add_patch(rect)
    plt.savefig('bounding_boxes.pdf', dpi =400)
    plt.show()

# This can be done before adjusting the bounding boxes incase you want to see the difference between the two.
#plot_bounding_boxes(image_path, bounding_boxes) 

# Plot the adjusted bounding boxes
#plot_bounding_boxes(image_path, adjusted_bounding_boxes)



# image_path = 'table_binarized.jpg'
# horizontal_guides, vertical_guides = parse_grid_file(file_path)
# plot_grid(image_path, horizontal_guides, vertical_guides)


