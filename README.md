# Drosophila Movement Tracker

A desktop app (PyQt5) for tracking multiple *Drosophila* flies, each in its own vial, from a live camera or a recorded video file — logging position, movement status, speed, and total distance to Excel, and exporting trajectory plots as PNGs.

---

## 1. Requirements

```bash
pip install PyQt5 opencv-python openpyxl matplotlib numpy
```

Optional, for extra features:

```bash
pip install pygrabber      # Windows only - shows real camera names (e.g. "Logitech BRIO")
                            # instead of generic "Camera 0", "Camera 1"...
```

**Run:**
```bash
python fly_tracker_gui.py # or the name by which one has downloaded
```

---

## 2. Quick Start

1. **Camera** section → pick your camera, resolution, and frame rate → **Start Cam**.
2. **Flies** section → set how many flies you have → **Apply**.
3. For each fly: select its row in the **Live Status** table, click **Set ROI**, then click-and-drag on the video feed to box in its vial. Set its real-world **Vial Length/Width (mm)**.
4. Adjust **Detection** sliders (Threshold, Min Size) until the crosshair marker reliably lands on the fly.
5. **Session** section → set a **Session Name** and **Log Interval** → **Start Tracking**.
6. When done, **Export** section → **Save Excel** and **Export Paths**.

---

## 3. Control Reference & Tips

### Source Mode
| Control | What it does |
|---|---|
| **Live Camera** / **Video File** | Switch between tracking from a live camera and batch-processing a pre-recorded video. |
| **Load Video** | Pick a video file (mp4/avi/mov/mkv) to analyze. |
| **Process Video** | Runs detection through the entire loaded video, showing a progress bar. Logging is paced by the *video's own timestamps* (frame ÷ its FPS), not wall-clock time — so a 1-second log interval always means 1 second of footage, regardless of how fast your PC processes the file. |

**Tip:** Video File mode reuses the exact same ROI, threshold, and export settings as Live mode — set your ROIs on the video's first frame the same way you would on a live feed.

### Camera
| Control | What it does |
|---|---|
| **Camera** | Auto-detected list of cameras actually connected to the system. Click **Refresh Cameras** after plugging one in or starting an OBS Virtual Camera. |
| **Resolution** | Requests a specific capture resolution (16:9 options listed first). |
| **Frame Rate** | Requests a specific FPS from the camera (15/24/30/60). |
| **Start Cam** / **Stop Cam** | Begin/end the live feed. |

**Tips:**
- If the picture opens in 4:3 when your camera supports 16:9, it's because a webcam's *default* mode is often 4:3 — the Resolution dropdown forces it into widescreen instead.
- If your requested resolution/FPS gets silently downgraded, the app will pop up a message telling you what the camera actually granted — that's a driver/hardware limit, not a setting you're missing.
- **If FPS is stuck low (e.g. 4–10fps) even at a "60fps-capable" resolution**, it's almost always a USB bandwidth issue: raw/uncompressed 1080p60 needs far more throughput than USB2.0 (and sometimes USB3.0) can sustain, so the camera silently drops to a slow raw mode. Try a lower resolution to confirm — if FPS jumps up there, that confirms it's a bandwidth ceiling. Using an **OBS Virtual Camera** as the source instead of the physical camera directly generally resolves this.

**Note:** For better frame-rate delivery and more reliable camera control (resolution/format negotiation), you can use an **OBS Virtual Camera** as the video source instead of feeding the physical camera directly — see the app's **OBS Integration** section.

### Flies
| Control | What it does |
|---|---|
| **Number of Flies** + **Apply** | Resize the fly list. Existing ROIs, names, and logged data are preserved for flies that still exist after resizing. |

**Tip:** Renaming is *not* done here — double-click the **Fly** name cell directly in the **Live Status** table instead.

### Vial ROI / Calibration
| Control | What it does |
|---|---|
| **Selected Fly** | Shows which fly the controls below apply to — determined by which row is selected in the Live Status table. |
| **Set ROI** | Arms click-and-drag mode; draw a box around the selected fly's vial on the video feed. |
| **Clear ROI** | Removes the selected fly's ROI. |
| **Vial Length / Width (mm)** | Real-world dimensions used to convert the ROI's pixels into millimeters. |
| **Grid Divisions** | Number of secondary reference grid lines drawn inside each ROI (0 = off). Purely visual, doesn't affect detection. |
| **Save Layout / Load Layout** | Save/restore all ROIs, names, vial sizes, and detection settings as a `.json` file — handy for reusing the same physical rig setup across sessions. |

**Tip:** Set all your ROIs and calibration once, then **Save Layout** — next session, **Load Layout** gets you back to a ready-to-track state in one click.

### Detection
| Control | What it does |
|---|---|
| **Threshold** | Brightness cutoff separating the dark fly from the white background (lower = only very dark pixels count as "fly"). |
| **Min Size** | Minimum blob area (in pixels) to count as a detection — filters out small noise/debris. |
| **Smoothing** | Moving-average window over recent positions, to reduce frame-to-frame jitter. |
| **Preview Mask** | Shows the black/white threshold mask instead of the color feed inside each ROI — useful for tuning Threshold visually. |

**Tips:**
- Turn on **Preview Mask** while adjusting **Threshold** — you want the fly to appear as a single solid white blob with minimal background noise.
- If the crosshair keeps jumping to dust or shadows, raise **Min Size** so tiny specks are filtered out.
- Too much **Smoothing** will lag behind fast, real fly movements — keep it low unless jitter is a real problem.

### Session
| Control | What it does |
|---|---|
| **Session Name** | Used to build export filenames (see Section 5). |
| **Log Interval (s)** | How often a data point is recorded (default 1s). |
| **Duration (min)** | Auto-stops tracking after this many minutes (0 = unlimited). |
| **Start Tracking / Stop Tracking** | Begin/end data logging (independent of whether the camera itself is running — ROI boxes and the crosshair marker are always visible on the feed regardless of tracking state). |
| **Reset Data** | Clears all logged data and total distance back to zero. |

### Export
| Control | What it does |
|---|---|
| **Export Folder** | Where output files are saved. |
| **Save Excel** | Saves the workbook — filename and columns detailed in Section 5. |
| **Export Paths** | Saves trajectory PNGs — two per fly (X-vs-time and Y-vs-time). |
| **Movement Threshold (mm)** | The cutoff used to classify a logged step as "idle" (0) vs "move" (1), and to decide whether it counts toward Total Distance. A step only counts as movement if the X change, Y change, or both are *strictly greater* than this value. |
| **Auto-Save** + **interval** | Periodically saves an Excel snapshot (`<session>_autosave.xlsx`) during a long tracking session, without needing to click Save Excel manually. |

**Tip:** A higher Movement Threshold makes the tracker more tolerant of camera/detection jitter being misread as "movement" — raise it if flies that are truly sitting still are showing up as "move" too often.

### Live Status table
Shows, per fly: current name (double-click to rename), whether it's currently detected, its live X/Y position, and its running **Total Distance (mm)**. Clicking a row also selects that fly for the ROI/Calibration controls.

---

## 4. Marker & Overlay Reference

- **Colored rectangle** — the ROI boundary for that fly (color cycles per fly).
- **Yellow crosshair (✕)** — the detected fly position for that frame. Always drawn whenever an ROI is set, independent of whether tracking/logging is active.
- **Thin gray grid lines** — the secondary reference grid inside each ROI (see Grid Divisions).

---

## 5. Export File Formats

### Excel workbook — `<session name>_<date>_<time>.xlsx`
One worksheet per fly (named after the fly). Columns:

| Column | Meaning |
|---|---|
| Timestamp | When this point was logged |
| X (mm) / Y (mm) | Position within the vial |
| Status (0=idle, 1=move) | `1` if the X or Y change since the previous point exceeded the Movement Threshold; `0` otherwise. Cell is colored red (idle) or green (move). |
| Instantaneous Speed (mm/s) | Raw distance moved since the previous point, divided by the elapsed time — computed for every row, regardless of idle/move status. |
| Avg Speed – Move Segment (mm/s) | For each continuous run of "move" (1) rows, the average speed across that whole segment (same value repeated on every row of the segment). Blank on idle rows. |
| Total Distance (mm) | Running cumulative distance — only accumulates on steps classified as "move" (see Movement Threshold). |

### Trajectory PNGs — `<session name>_<date>_<time>_<fly name>_X_vs_time.png` / `..._Y_vs_time.png`
Two plots per fly: X position and Y position, each against elapsed time in seconds from that fly's first logged point.

Filenames for both output types are generated automatically — there's no manual rename option, so exports stay consistently identifiable across sessions.

---

## 6. Known Limitations (Detection Algorithm)

The detector is a simple, fast, threshold-based approach — well suited to a controlled, high-contrast, white-background setup, but worth knowing its edges:

- **Static threshold** — doesn't adapt if lighting drifts over a long session.
- **"Largest blob wins"** — dust, smudges, or condensation on the vial can be mistaken for the fly if they're the largest dark region.
- **No shape validation** — any sufficiently large dark blob is accepted, not just fly-shaped ones.
- **No occlusion handling** — if the fly is briefly obscured, it's simply logged as "not detected" rather than estimated/interpolated.
- **Motion blur** at high fly speed can shrink the apparent blob and occasionally cause missed detections during fast moves.

If lighting consistency or debris becomes a real problem, an adaptive background-subtraction approach (e.g. `cv2.createBackgroundSubtractorMOG2`) would be a more robust upgrade than a fixed threshold.

---

## 7. Troubleshooting Quick Reference

| Symptom | Likely cause / fix |
|---|---|
| Feed opens at 4:3 | Camera defaulted to its narrow mode — pick a 16:9 option in **Resolution**. |
| FPS much lower than camera's rated max | USB bandwidth limit on raw video at high resolution/FPS — use an OBS Virtual Camera as the source instead. |
| Requested resolution/FPS not honored | Driver clamped to its nearest supported mode — the app will tell you what it actually got. |
| Crosshair jumps to wrong spot | Lower **Threshold**, raise **Min Size**, or enable **Preview Mask** to tune visually. |
| Fly shows "move" while sitting still | Raise the **Movement Threshold** slider. |
| Camera not in the list | Click **Refresh Cameras** — it only lists devices that are currently connected/on. |
