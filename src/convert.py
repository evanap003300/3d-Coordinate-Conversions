# Optitrack measurements:
# Height = 58 inches -> 1473.20 mm
# Width = 10 inches to the right -> 254.0 mm

# Fisheye camera:
# Height off the ground for the cameras is 5 1/2 + 5 inches for camera shaft 142.24 mm + 127.0 mm
# Camera Angle: 20 degrees 

# Image Dimentions: Width: 2160, Height: 2160
 
import numpy as np
import pandas as pd

# Store the final transformed coordinate data 
data_rows = []

def transform_coordinates(x, y, z, rotation_deg=(0, 0, 0), translation=[0, 0, 0]):
    rx, ry, rz = np.radians(rotation_deg)
    
    # Rotation matrices
    Rx = np.array([
        [1, 0, 0],
        [0, np.cos(rx), -np.sin(rx)],
        [0, np.sin(rx), np.cos(rx)]
    ])
    
    Ry = np.array([
        [np.cos(ry), 0, np.sin(ry)],
        [0, 1, 0],
        [-np.sin(ry), 0, np.cos(ry)]
    ])
    
    Rz = np.array([
        [np.cos(rz), -np.sin(rz), 0],
        [np.sin(rz), np.cos(rz), 0],
        [0, 0, 1]
    ])

    # Combined rotation matrix (ZYX order)
    R = Rz @ Ry @ Rx

    # Stack and transform
    coords = np.vstack((x, y, z))
    transformed = R @ coords + np.array(translation).reshape(3, 1)
    return transformed[0], transformed[1], transformed[2]

# Correction for horizontal offset (camera was placed 254 mm right of OptiTrack origin)
def fix_x_error(x):
    x_new = x - 254.0 
    return x_new

# Correction for vertical offset (OptiTrack origin is 1473.20 mm off the ground and the fisheye cameras are 142.24 mm + 127.0 mm off the ground)
def fix_y_error(y):
    y_new = y + 1473.20 - 142.24 - 127.0
    return y_new

# Nothing for now
def fix_z_error(z):
    return z

# Converts pixle coordinates to mm through using DPI conversion 
def convert_pixles_to_mm(x, y):
    new_x = x * (25.4/96)
    new_y = y * (25.4/96)
    return new_x, new_y

# Adds data to array for pandas excel output
def add_data(x, y, z):
    data_rows.append({
        "x": x,
        "y": y,
        "z": z
    })


# Exports the data to a excel file
def export_data():
    if not data_rows:
        print("No data to export")
        return
        
    df = pd.DataFrame(data_rows)
    output_filename = "../excel_files/OptiTrack.xlsx"
    df.to_excel(output_filename, index=False)
    print(f"xyz data saved to {output_filename}")
    print(f"Exported {len(data_rows)} 3D coordinate points")

def main():
    input_path = "../excel_files/input_coordinates.xlsx"
    df = pd.read_excel(input_path)

    img_width = 2160
    img_height = 2160

    for idx, row in df.iterrows():
        # Extract normalized 3D coordinates
        norm_x1 = row["x1"]
        norm_y1 = row["y1"]
        norm_x2 = row["x2"]
        norm_y2 = row["y2"]

        # Convert normalized to pixel values
        px1 = norm_x1 * img_width
        py1 = norm_y1 * img_height
        px2 = norm_x2 * img_width
        py2 = norm_y2 * img_height

        # Compute average position
        avg_x = (px1 + px2) / 2
        avg_y = (py1 + py2) / 2

        # Convert to mm
        x_mm, y_mm = convert_pixles_to_mm(avg_x, avg_y)
        z_mm, _ = convert_pixles_to_mm(px2, 0)  # Using right camera X as Z

        # Apply corrections
        x_corr = fix_x_error(x_mm)
        y_corr = fix_y_error(y_mm)
        z_corr = z_mm

        # Prepare for transformation
        x_arr = np.array([x_corr])
        y_arr = np.array([y_corr])
        z_arr = np.array([z_corr])

        x_t, y_t, z_t = transform_coordinates(
            x_arr, y_arr, z_arr,
            rotation_deg=(0, 20, 0),  # rotate up by 20°
            translation=[4000, -500, -700]  # scale for closer fit based on distances away from camera
        )

        # Final 3D point in OptiTrack frame
        final_x = z_t[0]
        final_y = y_t[0]
        final_z = x_t[0]

        add_data(final_x, final_y, final_z)

    export_data()

if __name__ == "__main__":
    main()
