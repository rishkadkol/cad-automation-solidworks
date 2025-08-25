import win32com.client
import pythoncom
import os
import csv

def create_cylinder(radius_mm, height_mm, save_path_part, save_path_stl):
    """Create a cylinder part in SolidWorks and export as STL"""
    pythoncom.CoInitialize()
    try:
        swApp = win32com.client.Dispatch("SldWorks.Application")
        swApp.Visible = True

        template_path = r"templates\Part.prtdot"  # Adjust path as needed
        if not os.path.exists(template_path):
            print(f"❌ Template not found: {template_path}")
            return False

        swApp.NewDocument(template_path, 0, 0, 0)
        model = swApp.ActiveDoc
        if model is None:
            print("❌ Could not open part document")
            return False

        # Convert mm → meters
        radius_m = float(radius_mm) / 1000.0
        height_m = float(height_mm) / 1000.0

        # Select Front Plane
        success = model.Extension.SelectByID2(
            "Front Plane", "PLANE", 0, 0, 0,
            False, 0, win32com.client.VARIANT(pythoncom.VT_DISPATCH, None), 0
        )
        if not success:
            print("❌ Could not select Front Plane")
            return False

        # Sketch circle
        model.SketchManager.InsertSketch(True)
        model.SketchManager.CreateCircle(0.0, 0.0, 0.0, radius_m, 0.0, 0.0)
        model.SketchManager.InsertSketch(True)

        # Extrude
        feat_mgr = model.FeatureManager
        feat = feat_mgr.FeatureExtrusion2(
            True, False, False,
            0, 0,
            height_m, 0.0,
            False, False, False, False,
            0.0, 0.0,
            False, False, False, False,
            True, False, False,
            0, 0, False
        )
        if feat is None:
            print("❌ Extrusion failed")
            return False

       

        # Save as SolidWorks part
        model.SaveAs(save_path_part)
        print(f"✅ Part saved: {save_path_part}")

        # Export as STL
        status = model.SaveAs3(save_path_stl, 0, 0)  # 0 = standard save, 0 = no options
        if status == 0:
            print(f"✅ STL exported: {save_path_stl}")
        else:
            print(f"⚠️ STL export may have failed (status: {status})")

        return True

    except Exception as e:
        print(f"❌ Exception: {e}")
        return False
    finally:
        pythoncom.CoUninitialize()


def batch_from_csv(csv_file, output_dir):
    """Read a CSV and generate multiple cylinder parts with STL export"""
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    with open(csv_file, newline='') as f:
        reader = csv.DictReader(f)
        for row in reader:
            name = row["part_name"]
            radius = row["radius_mm"]
            height = row["height_mm"]

            save_path_part = os.path.join(output_dir, f"{name}.SLDPRT")
            save_path_stl = os.path.join(output_dir, f"{name}.STL")

            print(f"\n🔹 Creating {name} (R={radius}mm, H={height}mm)...")
            create_cylinder(radius, height, save_path_part, save_path_stl)


if __name__ == "__main__":
    csv_file = r"CSV\cylinders.csv"  # Adjust path as needed
    output_dir = r"output"  # Adjust path as needed
    batch_from_csv(csv_file, output_dir)
