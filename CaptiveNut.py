# CaptiveNut.py
import FreeCAD as App
import FreeCADGui as Gui
from PySide2 import QtWidgets, QtCore, QtGui
import Part
import math
import os
import json

# Preset directory
PRESET_DIR = os.path.join(App.getUserMacroDir(), "CaptiveNut", "presets")
os.makedirs(PRESET_DIR, exist_ok=True)

# ===== INSERT DATABASE =====
INSERT_TYPES = {
    "M3 Captive Nut": {"nut": (5.5, 2.4), "bolt_clearance": 3.5, "pocket_style": "hex"},
    "M4 Captive Nut": {"nut": (7.0, 3.2), "bolt_clearance": 4.5, "pocket_style": "hex"},
    "M5 Captive Nut": {"nut": (8.0, 4.0), "bolt_clearance": 5.5, "pocket_style": "hex"},
    "M6 Captive Nut": {"nut": (10.0, 5.0), "bolt_clearance": 6.6, "pocket_style": "hex"},
    "M3 Heat-Set": {"nut": (4.0, 4.0), "bolt_clearance": 3.0, "pocket_style": "round", "pocket_diameter": 4.2},
    "M4 Heat-Set": {"nut": (5.0, 5.0), "bolt_clearance": 4.2, "pocket_style": "round", "pocket_diameter": 5.2},
    "M5 Heat-Set": {"nut": (6.0, 6.0), "bolt_clearance": 5.2, "pocket_style": "round", "pocket_diameter": 6.2},
    "M3 Square Nut": {"nut": (5.5, 5.5), "bolt_clearance": 3.5, "pocket_style": "rect", "thickness": 2.4},
    "M4 Square Nut": {"nut": (7.0, 7.0), "bolt_clearance": 4.5, "pocket_style": "rect", "thickness": 3.2},
    "M5 Square Nut": {"nut": (8.0, 8.0), "bolt_clearance": 5.5, "pocket_style": "rect", "thickness": 4.0},
}

PLACEMENT_MODES = ["Sketch Circles", "Import DXF", "Pattern Generator"]

# === GLOBAL WORKFLOW REGISTRY ===
_ACTIVE_WORKFLOWS = []

# === Helper: Get face subname like "Face3" ===
def get_face_subname(obj, face):
    for i, f in enumerate(obj.Shape.Faces):
        if f.isSame(face):
            return f"Face{i+1}"
    raise RuntimeError("Selected face not found in object")

# === Helper: Resolve to containing PartDesign::Body ===
def get_partdesign_body(obj):
    current = obj
    while current:
        if current.TypeId == "PartDesign::Body":
            return current
        if hasattr(current, 'InList') and current.InList:
            current = current.InList[0]
        else:
            break
    return None


class CaptiveNutDialog(QtWidgets.QDialog):
    def __init__(self, parent=None, face_center=None):
        super().__init__(parent)
        self.setWindowTitle("Captive Nut Placement")
        self.resize(400, 600)
        self.mode = "Sketch Circles"
        self.face_center = face_center or App.Vector(0, 0, 0)
        layout = QtWidgets.QVBoxLayout()

        preset_layout = QtWidgets.QHBoxLayout()
        self.preset_name = QtWidgets.QLineEdit()
        self.preset_name.setPlaceholderText("Preset name")
        self.save_preset_btn = QtWidgets.QPushButton("Save Preset")
        self.load_preset_btn = QtWidgets.QPushButton("Load Preset")
        self.save_preset_btn.clicked.connect(self.save_preset)
        self.load_preset_btn.clicked.connect(self.load_preset)
        preset_layout.addWidget(self.preset_name)
        preset_layout.addWidget(self.save_preset_btn)
        preset_layout.addWidget(self.load_preset_btn)
        layout.addLayout(preset_layout)

        mode_layout = QtWidgets.QFormLayout()
        self.mode_combo = QtWidgets.QComboBox()
        self.mode_combo.addItems(PLACEMENT_MODES)
        self.mode_combo.currentTextChanged.connect(self.on_mode_changed)
        mode_layout.addRow("Placement Mode:", self.mode_combo)
        layout.addLayout(mode_layout)

        self.pattern_group = QtWidgets.QGroupBox("Pattern Settings")
        pattern_layout = QtWidgets.QFormLayout()
        self.pattern_type = QtWidgets.QComboBox()
        self.pattern_type.addItems(["Linear", "Polar"])
        self.pattern_type.currentTextChanged.connect(self.on_pattern_type_changed)
        self.base_point = QtWidgets.QLineEdit(f"{self.face_center.x},{self.face_center.y},{self.face_center.z}")
        self.count_spin = QtWidgets.QSpinBox()
        self.count_spin.setRange(1, 100)
        self.count_spin.setValue(3)
        self.linear_spacing = QtWidgets.QDoubleSpinBox()
        self.linear_spacing.setRange(1, 200)
        self.linear_spacing.setValue(20)
        self.linear_spacing.setSuffix(" mm")
        self.polar_radius = QtWidgets.QDoubleSpinBox()
        self.polar_radius.setRange(1, 500)
        self.polar_radius.setValue(30)
        self.polar_radius.setSuffix(" mm")
        self.polar_angle = QtWidgets.QDoubleSpinBox()
        self.polar_angle.setRange(1, 360)
        self.polar_angle.setValue(30)
        self.polar_angle.setSuffix("°")

        pattern_layout.addRow("Pattern Type:", self.pattern_type)
        pattern_layout.addRow("Base Point (x,y,z):", self.base_point)
        pattern_layout.addRow("Count:", self.count_spin)
        pattern_layout.addRow("Spacing (Linear):", self.linear_spacing)
        pattern_layout.addRow("Radius (Polar):", self.polar_radius)
        pattern_layout.addRow("Angle Step (Polar):", self.polar_angle)
        self.pattern_group.setLayout(pattern_layout)
        self.pattern_group.setVisible(False)
        layout.addWidget(self.pattern_group)

        self.dxf_group = QtWidgets.QGroupBox("DXF Import")
        dxf_layout = QtWidgets.QFormLayout()
        self.dxf_file = QtWidgets.QLineEdit()
        self.dxf_browse = QtWidgets.QPushButton("Browse...")
        self.dxf_browse.clicked.connect(self.browse_dxf)
        dxf_file_layout = QtWidgets.QHBoxLayout()
        dxf_file_layout.addWidget(self.dxf_file)
        dxf_file_layout.addWidget(self.dxf_browse)
        dxf_layout.addRow("DXF File:", dxf_file_layout)
        self.dxf_layer_combo = QtWidgets.QComboBox()
        dxf_layout.addRow("Layer:", self.dxf_layer_combo)
        self.dxf_group.setLayout(dxf_layout)
        self.dxf_group.setVisible(False)
        layout.addWidget(self.dxf_group)

        common_group = QtWidgets.QGroupBox("Insert & Hole Settings")
        common_layout = QtWidgets.QFormLayout()
        self.type_combo = QtWidgets.QComboBox()
        self.type_combo.addItems(INSERT_TYPES.keys())
        self.pocket_depth = QtWidgets.QDoubleSpinBox()
        self.pocket_depth.setRange(1, 50)
        self.pocket_depth.setValue(6)
        self.pocket_depth.setSuffix(" mm")
        self.hole_type = QtWidgets.QComboBox()
        self.hole_type.addItems(["Through All", "Blind Hole", "Countersunk"])
        self.hole_type.currentTextChanged.connect(self.on_hole_type_changed)
        self.hole_depth = QtWidgets.QDoubleSpinBox()
        self.hole_depth.setRange(1, 100)
        self.hole_depth.setValue(20)
        self.hole_depth.setSuffix(" mm")
        self.countersink_angle = QtWidgets.QDoubleSpinBox()
        self.countersink_angle.setRange(60, 120)
        self.countersink_angle.setValue(90)
        self.countersink_angle.setSuffix("°")
        self.auto_boolean = QtWidgets.QCheckBox("Auto Boolean Cut (Part) / Auto Create (PartDesign)")
        self.auto_boolean.setChecked(True)
        common_layout.addRow("Insert Type:", self.type_combo)
        common_layout.addRow("Pocket Depth:", self.pocket_depth)
        common_layout.addRow("Hole Type:", self.hole_type)
        common_layout.addRow("Hole Depth:", self.hole_depth)
        common_layout.addRow("Countersink Angle:", self.countersink_angle)
        common_layout.addRow("", self.auto_boolean)
        common_group.setLayout(common_layout)
        layout.addWidget(common_group)

        self.on_hole_type_changed("Through All")
        self.on_mode_changed("Sketch Circles")
        self.on_pattern_type_changed("Linear")

        btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)
        self.setLayout(layout)

    def on_mode_changed(self, mode):
        self.mode = mode
        self.pattern_group.setVisible(mode == "Pattern Generator")
        self.dxf_group.setVisible(mode == "Import DXF")
        if mode == "Import DXF":
            self.update_dxf_layers()

    def on_pattern_type_changed(self, ptype):
        self.linear_spacing.setVisible(ptype == "Linear")
        self.polar_radius.setVisible(ptype == "Polar")
        self.polar_angle.setVisible(ptype == "Polar")

    def on_hole_type_changed(self, text):
        self.hole_depth.setEnabled(text != "Through All")
        self.countersink_angle.setEnabled(text == "Countersunk")

    def browse_dxf(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Select DXF", "", "DXF Files (*.dxf)")
        if path:
            self.dxf_file.setText(path)
            self.update_dxf_layers()

    def update_dxf_layers(self):
        self.dxf_layer_combo.clear()
        dxf_path = self.dxf_file.text()
        if not dxf_path or not os.path.isfile(dxf_path):
            return
        try:
            import importDXF
            obj = importDXF.open(dxf_path)
            layers = set()
            def collect_layers(obj):
                if hasattr(obj, 'Proxy') and hasattr(obj.Proxy, 'layer'):
                    layers.add(obj.Proxy.layer)
                for child in obj.OutList:
                    collect_layers(child)
            if obj:
                collect_layers(obj)
                App.ActiveDocument.removeObject(obj.Name)
            self.dxf_layer_combo.addItems(sorted(layers) if layers else ["All"])
        except Exception as e:
            App.Console.PrintError(f"Could not read DXF layers: {e}\n")
            self.dxf_layer_combo.addItem("All")

    def save_preset(self):
        name = self.preset_name.text().strip()
        if not name:
            QtWidgets.QMessageBox.warning(self, "Error", "Enter a preset name.")
            return
        preset = {
            "mode": self.mode_combo.currentText(),
            "insert_type": self.type_combo.currentText(),
            "pocket_depth": self.pocket_depth.value(),
            "hole_type": self.hole_type.currentText(),
            "hole_depth": self.hole_depth.value(),
            "countersink_angle": self.countersink_angle.value(),
            "auto_boolean": self.auto_boolean.isChecked(),
            "pattern_type": self.pattern_type.currentText(),
            "base_point": self.base_point.text(),
            "count": self.count_spin.value(),
            "linear_spacing": self.linear_spacing.value(),
            "polar_radius": self.polar_radius.value(),
            "polar_angle": self.polar_angle.value(),
            "dxf_file": self.dxf_file.text(),
        }
        path = os.path.join(PRESET_DIR, f"{name}.json")
        with open(path, 'w') as f:
            json.dump(preset, f, indent=2)
        App.Console.PrintMessage(f"✅ Preset saved: {name}\n")

    def load_preset(self):
        name = self.preset_name.text().strip()
        if not name:
            files = [f[:-5] for f in os.listdir(PRESET_DIR) if f.endswith('.json')]
            if not files:
                QtWidgets.QMessageBox.information(self, "No Presets", "No presets found.")
                return
            name, ok = QtWidgets.QInputDialog.getItem(self, "Load Preset", "Select preset:", files, 0, False)
            if not ok:
                return
        path = os.path.join(PRESET_DIR, f"{name}.json")
        if not os.path.isfile(path):
            QtWidgets.QMessageBox.critical(self, "Error", f"Preset not found: {name}")
            return
        with open(path, 'r') as f:
            preset = json.load(f)
        self.mode_combo.setCurrentText(preset.get("mode", "Sketch Circles"))
        self.type_combo.setCurrentText(preset.get("insert_type", "M3 Captive Nut"))
        self.pocket_depth.setValue(preset.get("pocket_depth", 6))
        self.hole_type.setCurrentText(preset.get("hole_type", "Through All"))
        self.hole_depth.setValue(preset.get("hole_depth", 20))
        self.countersink_angle.setValue(preset.get("countersink_angle", 90))
        self.auto_boolean.setChecked(preset.get("auto_boolean", True))
        self.pattern_type.setCurrentText(preset.get("pattern_type", "Linear"))
        self.base_point.setText(preset.get("base_point", "0,0,0"))
        self.count_spin.setValue(preset.get("count", 3))
        self.linear_spacing.setValue(preset.get("linear_spacing", 20))
        self.polar_radius.setValue(preset.get("polar_radius", 30))
        self.polar_angle.setValue(preset.get("polar_angle", 30))
        self.dxf_file.setText(preset.get("dxf_file", ""))
        self.preset_name.setText(name)
        self.on_mode_changed(self.mode_combo.currentText())
        self.on_pattern_type_changed(self.pattern_type.currentText())
        self.on_hole_type_changed(self.hole_type.currentText())
        App.Console.PrintMessage(f"✅ Preset loaded: {name}\n")


class CaptiveNutCommand:
    def GetResources(self):
        return {
            'Pixmap': 'CaptiveNut.svg',
            'MenuText': 'Place Captive Nuts',
            'ToolTip': 'Place nut pockets & bolt holes via sketch, DXF, or pattern (supports Part & PartDesign)'
        }

    def IsActive(self):
        return True

    def Activated(self):
        sel = Gui.Selection.getSelectionEx()
        if not sel:
            QtWidgets.QMessageBox.warning(Gui.getMainWindow(), "Error", "Select a planar face.")
            return
        face = sel[0].SubObjects[0]
        obj = sel[0].Object

        if not (hasattr(face, 'Surface') and isinstance(face.Surface, Part.Plane)):
            QtWidgets.QMessageBox.warning(Gui.getMainWindow(), "Error", "Only planar faces supported.")
            return

        body = get_partdesign_body(obj)
        is_partdesign = body is not None

        face_center = face.CenterOfMass
        dialog = CaptiveNutDialog(Gui.getMainWindow(), face_center)
        if not dialog.exec_():
            return

        dialog_data = {
            "mode": dialog.mode_combo.currentText(),
            "insert_type": dialog.type_combo.currentText(),
            "pocket_depth": dialog.pocket_depth.value(),
            "hole_type": dialog.hole_type.currentText(),
            "hole_depth": dialog.hole_depth.value(),
            "countersink_angle": dialog.countersink_angle.value(),
            "auto_boolean": dialog.auto_boolean.isChecked(),
            "pattern_type": dialog.pattern_type.currentText(),
            "base_point": dialog.base_point.text(),
            "count": dialog.count_spin.value(),
            "linear_spacing": dialog.linear_spacing.value(),
            "polar_radius": dialog.polar_radius.value(),
            "polar_angle": dialog.polar_angle.value(),
            "dxf_file": dialog.dxf_file.text(),
            "dxf_layer": dialog.dxf_layer_combo.currentText(),
        }

        centers = []
        if dialog_data["mode"] == "Sketch Circles":
            App.Console.PrintMessage("👉 Please sketch circles and close the sketch.\n")
            workflow = CaptiveNutSketchWorkflow(obj, face, body, dialog_data, is_partdesign)
            return
        elif dialog_data["mode"] == "Import DXF":
            dxf_path = dialog_data["dxf_file"]
            if not dxf_path or not os.path.isfile(dxf_path):
                QtWidgets.QMessageBox.critical(Gui.getMainWindow(), "Error", "Invalid DXF file.")
                return
            centers = self.extract_centers_from_dxf(dxf_path, dialog_data["dxf_layer"])
        elif dialog_data["mode"] == "Pattern Generator":
            try:
                base_str = dialog_data["base_point"].replace(',', ' ').split()
                base = App.Vector(float(base_str[0]), float(base_str[1]), float(base_str[2]))
                count = dialog_data["count"]
                if dialog_data["pattern_type"] == "Linear":
                    spacing = dialog_data["linear_spacing"]
                    centers = [base + App.Vector(i * spacing, 0, 0) for i in range(count)]
                else:
                    radius = dialog_data["polar_radius"]
                    angle_step = dialog_data["polar_angle"]
                    centers = []
                    for i in range(count):
                        ang = math.radians(i * angle_step)
                        x = base.x + radius * math.cos(ang)
                        y = base.y + radius * math.sin(ang)
                        centers.append(App.Vector(x, y, base.z))
            except Exception as e:
                QtWidgets.QMessageBox.critical(Gui.getMainWindow(), "Error", f"Invalid pattern input:\n{e}")
                return

        if not centers:
            QtWidgets.QMessageBox.warning(Gui.getMainWindow(), "No Centers", "No placement points found.")
            return

        if is_partdesign:
            if dialog_data["auto_boolean"]:
                reply = QtWidgets.QMessageBox.question(
                    Gui.getMainWindow(), "Create Pockets in Body?",
                    f"Add {len(centers)} captive nut pockets directly into the PartDesign Body?",
                    QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No
                )
                if reply == QtWidgets.QMessageBox.Yes:
                    for i, center in enumerate(centers):
                        try:
                            self.create_partdesign_pocket(
                                body, face, center,
                                INSERT_TYPES[dialog_data["insert_type"]],
                                dialog_data["pocket_depth"],
                                dialog_data["hole_type"],
                                dialog_data["hole_depth"],
                                dialog_data["countersink_angle"],
                                sketch=None
                            )
                        except Exception as e:
                            App.Console.PrintError(f"Failed pocket #{i+1}: {e}\n")
                    App.Console.PrintMessage(f"✅ {len(centers)} pockets added to Body.\n")
            else:
                App.Console.PrintMessage(f"ℹ️ Ready to create {len(centers)} pockets in Body (auto-create disabled).\n")
        else:
            preview = self.make_preview(centers, obj, face, dialog_data)
            if not preview:
                return

            doc = App.ActiveDocument
            preview_obj = doc.addObject("Part::Feature", "CaptiveNut_Preview")
            preview_obj.Shape = preview
            preview_obj.ViewObject.Transparency = 70
            preview_obj.ViewObject.ShapeColor = (0.0, 0.8, 1.0)
            doc.recompute()

            if dialog_data["auto_boolean"]:
                reply = QtWidgets.QMessageBox.question(
                    Gui.getMainWindow(), "Confirm Boolean Cut",
                    f"Place {len(centers)} inserts and cut from body?\n(Preview shown in cyan)",
                    QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No
                )
                if reply == QtWidgets.QMessageBox.Yes:
                    try:
                        cut_obj = doc.addObject("Part::Cut", "CaptiveNut_Cut")
                        cut_obj.Base = obj
                        cut_obj.Tool = preview_obj
                        doc.recompute()
                        App.Console.PrintMessage(f"✅ {len(centers)} inserts placed and cut.\n")
                    except Exception as e:
                        QtWidgets.QMessageBox.critical(Gui.getMainWindow(), "Boolean Failed", str(e))
            else:
                App.Console.PrintMessage("ℹ️ Preview created. Use Boolean Cut manually if desired.\n")

    # =============== PARTDESIGN SUPPORT ===============
    def create_partdesign_pocket(self, body, base_face, center, insert, pocket_depth, hole_type, hole_depth, countersink_angle, sketch=None):
        doc = App.ActiveDocument
        
        # === STEP 1: NUT POCKET (depth = pocket_depth) ===
        nut_sketch = doc.addObject("Sketcher::SketchObject", "CaptiveNut_Nut")
        if sketch is not None:
            # Clone attachment from user sketch
            nut_sketch.AttachmentSupport = sketch.AttachmentSupport
            nut_sketch.MapMode = sketch.MapMode
            nut_sketch.Placement = sketch.Placement
        else:
            # Create new attachment
            face_name = get_face_subname(body, base_face)
            nut_sketch.AttachmentSupport = [(body, face_name)]
            nut_sketch.MapMode = 'FlatFace'
        doc.recompute()
        
        local_center = nut_sketch.Placement.inverse().multVec(center)
        
        # Draw nut pocket
        style = insert["pocket_style"]
        if style == "hex":
            w = insert["nut"][0]
            outer = w + 1.0
            r = outer / 2
            for i in range(6):
                ang1 = math.radians(60 * i + 30)
                ang2 = math.radians(60 * (i + 1) + 30)
                p1 = App.Vector(local_center.x + r * math.cos(ang1), local_center.y + r * math.sin(ang1))
                p2 = App.Vector(local_center.x + r * math.cos(ang2), local_center.y + r * math.sin(ang2))
                nut_sketch.addGeometry(Part.LineSegment(App.Vector(p1.x, p1.y, 0), App.Vector(p2.x, p2.y, 0)), False)
        elif style == "rect":
            w, h = insert["nut"]
            clearance = 0.5
            pts = [
                App.Vector(local_center.x - w/2 - clearance, local_center.y - h/2 - clearance),
                App.Vector(local_center.x + w/2 + clearance, local_center.y - h/2 - clearance),
                App.Vector(local_center.x + w/2 + clearance, local_center.y + h/2 + clearance),
                App.Vector(local_center.x - w/2 - clearance, local_center.y + h/2 + clearance),
                App.Vector(local_center.x - w/2 - clearance, local_center.y - h/2 - clearance)
            ]
            for i in range(4):
                nut_sketch.addGeometry(Part.LineSegment(pts[i], pts[i+1]), False)
        else:  # round
            d = insert.get("pocket_diameter", insert["nut"][0] + 0.2)
            nut_sketch.addGeometry(Part.Circle(App.Vector(local_center.x, local_center.y, 0), App.Vector(0,0,1), d/2), False)
        
        doc.recompute()
        
        # Create nut pocket feature
        nut_pocket = doc.addObject("PartDesign::Pocket", "CaptiveNut_NutPocket")
        nut_pocket.Profile = nut_sketch
        nut_pocket.Midplane = 0
        nut_pocket.Reversed = 0
        nut_pocket.Type = 0  # Blind
        nut_pocket.Length = pocket_depth
        
        body.addObject(nut_sketch)
        body.addObject(nut_pocket)
        doc.recompute()
        
        # === STEP 2: BOLT HOLE ===
        bolt_sketch = doc.addObject("Sketcher::SketchObject", "CaptiveNut_Bolt")
        if sketch is not None:
            bolt_sketch.AttachmentSupport = sketch.AttachmentSupport
            bolt_sketch.MapMode = sketch.MapMode
            bolt_sketch.Placement = sketch.Placement
        else:
            face_name = get_face_subname(body, base_face)
            bolt_sketch.AttachmentSupport = [(body, face_name)]
            bolt_sketch.MapMode = 'FlatFace'
        doc.recompute()
        
        bolt_center = bolt_sketch.Placement.inverse().multVec(center)
        bolt_r = insert["bolt_clearance"] / 2
        bolt_sketch.addGeometry(Part.Circle(App.Vector(bolt_center.x, bolt_center.y, 0), App.Vector(0,0,1), bolt_r), False)
        doc.recompute()
        
        # Create bolt hole feature
        bolt_hole = doc.addObject("PartDesign::Pocket", "CaptiveNut_BoltHole")
        bolt_hole.Profile = bolt_sketch
        bolt_hole.Midplane = 0
        bolt_hole.Reversed = 0
        
        if hole_type == "Through All":
            bolt_hole.Type = 1  # ThroughAll
        elif hole_type == "Countersunk":
            QtWidgets.QMessageBox.warning(
                Gui.getMainWindow(),
                "Countersink Not Supported",
                "PartDesign Pockets do not support countersinks.\nUsing blind hole."
            )
            bolt_hole.Type = 0
            bolt_hole.Length = hole_depth
        else:  # Blind
            bolt_hole.Type = 0
            bolt_hole.Length = hole_depth
        
        body.addObject(bolt_sketch)
        body.addObject(bolt_hole)
        doc.recompute()
        return nut_pocket, bolt_hole

    # =============== SKETCH WORKFLOW HELPERS ===============
    def extract_centers_from_sketch(self, sketch):
        centers = []
        for geo in sketch.Geometry:
            if hasattr(geo, 'Radius') and hasattr(geo, 'Center'):
                centers.append(App.Vector(geo.Center.x, geo.Center.y, 0))
        if not centers:
            return []
        mat = sketch.Placement.toMatrix()
        return [mat.multiply(c) for c in centers]

    def extract_centers_from_dxf(self, dxf_path, layer_filter):
        centers = []
        try:
            import importDXF
            dxf_obj = importDXF.open(dxf_path)
            if not dxf_obj:
                return centers
            def traverse(obj):
                results = []
                if hasattr(obj, 'Proxy') and layer_filter != "All":
                    if getattr(obj.Proxy, 'layer', None) != layer_filter:
                        return results
                if hasattr(obj, 'Shape'):
                    for edge in obj.Shape.Edges:
                        if hasattr(edge.Curve, 'Radius'):
                            results.append(edge.Curve.Center)
                for child in obj.OutList:
                    results.extend(traverse(child))
                return results
            centers = traverse(dxf_obj)
            App.ActiveDocument.removeObject(dxf_obj.Name)
        except Exception as e:
            App.Console.PrintError(f"DXF processing error: {e}\n")
        return centers

    def make_preview(self, centers, base_obj, face, dialog_data):
        insert = INSERT_TYPES[dialog_data["insert_type"]]
        pocket_depth = dialog_data["pocket_depth"]
        hole_type = dialog_data["hole_type"]
        hole_depth = dialog_data["hole_depth"]
        countersink_angle = dialog_data["countersink_angle"]
        normal = face.normalAt(0, 0)
        com = base_obj.Shape.CenterOfMass
        face_center = face.CenterOfMass
        if normal.dot(face_center - com) > 0:
            normal = normal * -1

        shapes = []
        for center in centers:
            placement = App.Placement()
            placement.Base = center
            placement.Rotation = App.Rotation(App.Vector(0, 0, 1), normal)
            pocket = self.make_pocket(insert, pocket_depth)
            pocket.Placement = placement
            if hole_type == "Through All":
                hole = self.make_bolt_hole(insert, 1000)
            elif hole_type == "Countersunk":
                hole = self.make_countersunk_hole(insert, hole_depth, countersink_angle)
            else:
                hole = self.make_bolt_hole(insert, hole_depth)
            hole.Placement = placement
            shapes.extend([pocket, hole])
        return Part.makeCompound(shapes)

    def make_pocket(self, data, depth):
        style = data["pocket_style"]
        if style == "hex":
            w = data["nut"][0]
            outer = w + 1.0
            pts = [App.Vector((outer/2)*math.cos(math.radians(60*i+30)), 
                              (outer/2)*math.sin(math.radians(60*i+30)), 0) for i in range(6)]
            pts.append(pts[0])
            return Part.Face(Part.makePolygon(pts)).extrude(App.Vector(0,0,-depth))
        elif style == "rect":
            w, h = data["nut"]
            clearance = 0.5
            pts = [
                App.Vector(-w/2-clearance, -h/2-clearance, 0),
                App.Vector(w/2+clearance, -h/2-clearance, 0),
                App.Vector(w/2+clearance, h/2+clearance, 0),
                App.Vector(-w/2-clearance, h/2+clearance, 0),
                App.Vector(-w/2-clearance, -h/2-clearance, 0)
            ]
            return Part.Face(Part.makePolygon(pts)).extrude(App.Vector(0,0,-depth))
        else:
            d = data.get("pocket_diameter", data["nut"][0] + 0.2)
            return Part.makeCylinder(d/2, depth, App.Vector(0,0,0), App.Vector(0,0,-1))

    def make_bolt_hole(self, data, depth):
        return Part.makeCylinder(data["bolt_clearance"]/2, depth, App.Vector(0,0,0), App.Vector(0,0,-1))

    def make_countersunk_hole(self, data, depth, angle):
        d_clear = data["bolt_clearance"]
        d_head = d_clear * 2.0
        angle_rad = math.radians(angle / 2)
        head_depth = (d_head - d_clear) / (2 * math.tan(angle_rad))
        bottom = Part.makeCylinder(d_clear/2, depth, App.Vector(0,0,0), App.Vector(0,0,-1))
        cone = Part.makeCone(d_head/2, d_clear/2, head_depth, App.Vector(0,0,0), App.Vector(0,0,-1))
        return bottom.fuse(cone)


class CaptiveNutSketchWorkflow:
    def __init__(self, base_obj, face, body, dialog_data, is_partdesign):
        self.base_obj = base_obj
        self.face = face
        self.body = body
        self.dialog_data = dialog_data
        self.is_partdesign = is_partdesign
        self.doc = App.ActiveDocument
        self.sketch = None
        self.sketch_name = None
        self.timer = None
        self.setup_sketch()
        _ACTIVE_WORKFLOWS.append(self)

    def setup_sketch(self):
        self.sketch = self.doc.addObject("Sketcher::SketchObject", "CaptiveNut_UserSketch")
        face_name = get_face_subname(self.base_obj, self.face)
        self.sketch.AttachmentSupport = [(self.base_obj, face_name)]
        self.sketch.MapMode = 'FlatFace'
        self.doc.recompute()
        
        if self.is_partdesign:
            self.body.addObject(self.sketch)
        
        self.sketch_name = self.sketch.Name
        Gui.ActiveDocument.setEdit(self.sketch.Name)
        
        self.timer = QtCore.QTimer()
        self.timer.setSingleShot(True)
        self.timer.timeout.connect(self.start_polling)
        self.timer.start(1500)

    def start_polling(self):
        self.timer = QtCore.QTimer()
        self.timer.timeout.connect(self.check_sketch_closed)
        self.timer.start(1000)

    def check_sketch_closed(self):
        in_edit = Gui.ActiveDocument.getInEdit()
        if in_edit is None:
            self.timer.stop()
            self.cleanup()
            self.on_sketch_finished()

    def cleanup(self):
        if self in _ACTIVE_WORKFLOWS:
            _ACTIVE_WORKFLOWS.remove(self)
        if self.timer and self.timer.isActive():
            self.timer.stop()

    def on_sketch_finished(self):
        try:
            centers = CaptiveNutCommand().extract_centers_from_sketch(self.sketch)
            if not centers:
                QtWidgets.QMessageBox.warning(Gui.getMainWindow(), "No Circles", "No circles found in sketch.")
                return

            if self.is_partdesign:
                if self.dialog_data["auto_boolean"]:
                    reply = QtWidgets.QMessageBox.question(
                        Gui.getMainWindow(), "Create Pockets?",
                        f"Add {len(centers)} pockets to Body from sketch?",
                        QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No
                    )
                    if reply == QtWidgets.QMessageBox.Yes:
                        for i, center in enumerate(centers):
                            try:
                                # ✅ Pass user's sketch to create aligned, separate nut + bolt features
                                CaptiveNutCommand().create_partdesign_pocket(
                                    self.body, self.face, center,
                                    INSERT_TYPES[self.dialog_data["insert_type"]],
                                    self.dialog_data["pocket_depth"],
                                    self.dialog_data["hole_type"],
                                    self.dialog_data["hole_depth"],
                                    self.dialog_data["countersink_angle"],
                                    sketch=self.sketch
                                )
                            except Exception as e:
                                App.Console.PrintError(f"Failed pocket #{i+1}: {e}\n")
                        App.Console.PrintMessage(f"✅ {len(centers)} pockets added.\n")
                else:
                    App.Console.PrintMessage(f"ℹ️ Sketch processed. Auto-create disabled.\n")
            else:
                preview = CaptiveNutCommand().make_preview(centers, self.base_obj, self.face, self.dialog_data)
                if not preview:
                    return
                preview_obj = self.doc.addObject("Part::Feature", "CaptiveNut_Preview")
                preview_obj.Shape = preview
                preview_obj.ViewObject.Transparency = 70
                preview_obj.ViewObject.ShapeColor = (0.0, 0.8, 1.0)
                self.doc.recompute()

                if self.dialog_data["auto_boolean"]:
                    reply = QtWidgets.QMessageBox.question(
                        Gui.getMainWindow(), "Confirm Boolean Cut",
                        f"Cut {len(centers)} inserts from solid?",
                        QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No
                    )
                    if reply == QtWidgets.QMessageBox.Yes:
                        cut_obj = self.doc.addObject("Part::Cut", "CaptiveNut_Cut")
                        cut_obj.Base = self.base_obj
                        cut_obj.Tool = preview_obj
                        self.doc.recompute()
                        App.Console.PrintMessage(f"✅ Boolean cut performed.\n")
                else:
                    App.Console.PrintMessage("ℹ️ Preview created from sketch.\n")
        except Exception as e:
            App.Console.PrintError(f"Sketch processing error: {e}\n")
            import traceback
            traceback.print_exc()
        finally:
            self.cleanup()


Gui.addCommand("CaptiveNut_Command", CaptiveNutCommand())