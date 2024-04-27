import os
from selenium.webdriver.chrome.service import Service
from selenium import webdriver
from bimedis import bimedis
import pandas as pd
from surplex import surplex


def setup_driver(chromedriver_path):
    service = Service(executable_path=chromedriver_path)
    return webdriver.Chrome(service=service)


def main():
    # Define paths
    chromedriver_path = r"\\SERVER2022-DC\Workload_Data\pythonProject\chromedriver.exe"
    output_file = 'OUTPUT.xlsx'
    file_path = 'manufacture.xlsx'

    # Read Excel file
    if not os.path.exists(file_path):
        print(f"File {file_path} not found.")
        return
    df = pd.read_excel(file_path)

    # Setup WebDriver
    driver = setup_driver(chromedriver_path)

    try:
        # Iterate over rows in DataFrame
        for index, row in df.iterrows():
            search_term = row["שמות יצרני ציוד רפואי"]
            surplex()
            bimedis()
    finally:
        driver.quit()


def clean_data_and_remove_images(file_path, images_folder_path):
    # Read Excel file

    df = pd.read_excel(file_path)

    # Filter out rows where 'PRICE' column has 'Not available', 'unknown', or is empty
    filtered_df = df[
        ~df['price'].astype(str).str.contains('Not available|unknown') &  # checks for 'Not available' or 'unknown'
        df['price'].notna() &  # ensures that the PRICE column is not NaN
        df['price'].ne('')]  # ensures that the PRICE column is not an empty string

    # Find the difference in indices to identify which rows got removed
    removed_indices = set(df.index) - set(filtered_df.index)

    # List for deleted images
    images_to_delete = []

    # Check for images corresponding to the removed rows
    for index in removed_indices:
        image_id = df.at[index, 'ID']  # Assuming the image identifier column is named 'ID'
        image_path = os.path.join(images_folder_path, f"{image_id}.png")
        if os.path.exists(image_path):
            os.remove(image_path)
            images_to_delete.append(image_path)
            print(f"Deleted image: {image_path}")
        else:
            print(f"Image file does not exist: {image_path}")

    # Save the filtered DataFrame to a new Excel file
    filtered_file_path = 'OUTPUT.xlsx'
    filtered_df.to_excel(filtered_file_path, index=False)

    return filtered_file_path, images_to_delete


if __name__ == "__main__":
    # Define paths
    output_file = 'OUTPUT.xlsx'
    images_folder_path = f"\\\\SERVER2022-DC\\Workload_Data\\pythonProject\\screenshots"
    # Run the function
    cleaned_file, deleted_images = clean_data_and_remove_images(output_file, images_folder_path)

    print(f"Cleaned data saved to {cleaned_file}.")
    print("Deleted images:", deleted_images)
    main()
