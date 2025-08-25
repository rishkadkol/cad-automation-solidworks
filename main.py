from generate_model import create_cylinder
import os

def main():
    # Get user input
    part_name = input("Enter part name (without extension): ").strip()
    radius = float(input("Enter radius in mm: "))
    height = float(input("Enter height in mm: "))

    # Set up save path
    save_dir = r"file_path"  # Replace with your desired directory
    os.makedirs(save_dir, exist_ok=True)
    save_path = os.path.join(save_dir, f"{part_name}.SLDPRT")

    # Call the function
    create_cylinder(radius, height, save_path)

if __name__ == "__main__":
    main()