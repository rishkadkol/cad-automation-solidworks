import win32com.client
import pythoncom
import os

def create_cylinder(radius_mm, height_mm, save_path_part):
    pythoncom.CoInitialize()
    try:
        swApp = win32com.client.Dispatch("SldWorks.Application")
        swApp.Visible = True

        template_path = r"template_path\Part.prtdot"  # Update with your actual template path
        if not os.path.exists(template_path):
            print(f"❌ Template not found: {template_path}")
            return

        swApp.NewDocument(template_path, 0, 0, 0)
        model = swApp.ActiveDoc
        if model is None:
            print("❌ Could not open part document")
            return

        # ✅ Convert mm to meters for SolidWorks API
        radius_m = float(radius_mm) / 1000.0
        height_m = float(height_mm) / 1000.0

        # Select Front Plane
        print("📐 Selecting Front Plane...")
        success = model.Extension.SelectByID2(
            "Front Plane", "PLANE", 0, 0, 0,
            False, 0, win32com.client.VARIANT(pythoncom.VT_DISPATCH, None), 0
        )
        print(f"✅ Plane selected: {success}")
        # Start Sketch
        print("✏️ Creating Sketch...")
        model.SketchManager.InsertSketch(True)
        model.SketchManager.CreateCircle(0.0, 0.0, 0.0, radius_m, 0.0, 0.0)
        model.SketchManager.InsertSketch(True)  # End sketch

        # Extrude Feature
        print("🛠 Extruding...")
        feat_mgr = model.FeatureManager
        feat_mgr.FeatureExtrusion2(
            True,       # Solid feature
            False,      # Thin
            False,      # Direction 1 Draft
            0,          # End condition 1: Blind
            0,          # End condition 2
            height_m,   # Depth1
            0.0,        # Depth2
            False, False, False, False,
            0.0, 0.0,
            False, False, False, False,
            True,       # Merge result
            False, False,
            0, 0, False
        )

        

        # Save Part
        print("💾 Saving part...")
        model.SaveAs(save_path_part)
        print(f"✅ Part saved at: {save_path_part}")

    except Exception as e:
        print(f"❌ Exception: {e}")
    finally:
        pythoncom.CoUninitialize()



