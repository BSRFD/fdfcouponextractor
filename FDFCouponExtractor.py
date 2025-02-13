import fitz  # PyMuPDF
import os
import glob
import json
import io
import hashlib
import sys
from PIL import Image
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

CONFIG_FILE = "config.json"
Image.MAX_IMAGE_PIXELS = None

def save_config(source, destination, extract_images, convert_pdf, 
               suppress_messages, delete_fdf, coupons_per_page, merge_pdf):
    config_data = {
        "source_dir": source,
        "destination_dir": destination,
        "extract_images": extract_images,
        "convert_pdf": convert_pdf,
        "suppress_messages": suppress_messages,
        "delete_fdf": delete_fdf,
        "coupons_per_page": coupons_per_page,
        "merge_pdf": merge_pdf
    }
    with open(CONFIG_FILE, "w") as f:
        json.dump(config_data, f)

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    return None

def select_folder(entry_field):
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        entry_field.delete(0, tk.END)
        entry_field.insert(0, folder_selected)

def save_and_close(root, source_entry, dest_entry, extract_var, pdf_var, 
                   suppress_var, delete_var, coupons_var, merge_var):
    source = source_entry.get()
    destination = dest_entry.get()
    extract_images = extract_var.get()
    convert_pdf = pdf_var.get()
    suppress_messages = suppress_var.get()
    delete_fdf = delete_var.get()
    coupons_per_page = coupons_var.get()
    merge_pdf = merge_var.get()

    if not source or not destination:
        messagebox.showerror("Error", "Both source and destination folders must be selected!")
        return

    save_config(source, destination, extract_images, convert_pdf, 
                suppress_messages, delete_fdf, coupons_per_page, merge_pdf)
    root.destroy()

def get_user_settings():
    root = tk.Tk()
    root.title("Setup Configuration")
    window_width = 520
    window_height = 320
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")
    root.resizable(False, False)
    root.grid_columnconfigure(1, weight=1)

    tk.Label(root, text="Source:").grid(row=0, column=0, padx=10, pady=3, sticky="w")
    source_entry = tk.Entry(root, width=45)
    source_entry.grid(row=0, column=1, padx=10, pady=3, sticky="ew")
    tk.Button(root, text="Browse", command=lambda: select_folder(source_entry)).grid(row=0, column=2, padx=5, pady=3)

    tk.Label(root, text="Destination:").grid(row=1, column=0, padx=10, pady=3, sticky="w")
    dest_entry = tk.Entry(root, width=45)
    dest_entry.grid(row=1, column=1, padx=10, pady=3, sticky="ew")
    tk.Button(root, text="Browse", command=lambda: select_folder(dest_entry)).grid(row=1, column=2, padx=5, pady=3)

    check_frame = tk.Frame(root)
    check_frame.grid(row=2, column=0, columnspan=3, padx=10, pady=5, sticky="w")

    extract_var = tk.BooleanVar(value=True)
    tk.Checkbutton(check_frame, text="Extract coupon images [PNG]", variable=extract_var).pack(anchor="w", pady=2)

    pdf_var = tk.BooleanVar(value=True)
    tk.Checkbutton(check_frame, text="Create .PDF file(s)", variable=pdf_var).pack(anchor="w", pady=2)

    merge_var = tk.BooleanVar(value=False)
    tk.Checkbutton(check_frame, text="Combine multiple .FDF into single .PDF", variable=merge_var).pack(anchor="w", pady=2)

    delete_var = tk.BooleanVar(value=False)
    tk.Checkbutton(check_frame, text="Delete .FDF file(s) after processing", 
                  variable=delete_var).pack(anchor="w", pady=2)

    suppress_var = tk.BooleanVar(value=False)
    tk.Checkbutton(check_frame, text="Hide successful completion message", 
                  variable=suppress_var).pack(anchor="w", pady=2)

    # Coupons per page selector
    coupons_frame = tk.Frame(root)
    coupons_frame.grid(row=3, column=0, columnspan=3, padx=10, pady=5, sticky="w")
    tk.Label(coupons_frame, text="Coupons per page:").pack(side=tk.LEFT)
    coupons_var = tk.StringVar(value="4")
    coupons_menu = ttk.Combobox(coupons_frame, textvariable=coupons_var, 
                               values=["1", "2", "3", "4", "5"], width=3, state="readonly")
    coupons_menu.pack(side=tk.LEFT, padx=5)

    tk.Button(root, text="Save & Continue", 
             command=lambda: save_and_close(root, source_entry, dest_entry, 
                                           extract_var, pdf_var, suppress_var, 
                                           delete_var, coupons_var, merge_var)
             ).grid(row=4, column=1, pady=10)

    return root

def get_image_hash(image_bytes):
    return hashlib.sha256(image_bytes).hexdigest()

def optimize_image_storage(img_bytes, ext):
    """Optimize storage without recompressing lossy formats"""
    if ext.lower() in ('jpg', 'jpeg'):
        return img_bytes
    try:
        img = Image.open(io.BytesIO(img_bytes))
        if img.mode == 'RGBA':
            img = img.convert('RGB')
        optimized = io.BytesIO()
        img.save(optimized, format=ext, quality=100, optimize=True)
        return optimized.getvalue()
    except:
        return img_bytes

def create_pdf(output_path, images, coupons_per_page):
    """Create PDF with images arranged according to coupons_per_page setting"""
    pdf_doc = fitz.open()
    a4_width = 595
    a4_height = 842
    vertical_gap = 14
    horizontal_margin = 20

    cpp = max(1, min(5, coupons_per_page))
    num_gaps = cpp - 1
    section_height = (a4_height - (num_gaps * vertical_gap)) / cpp
    
    for i in range(0, len(images), cpp):
        batch = images[i:i+cpp]
        page = pdf_doc.new_page(width=a4_width, height=a4_height)
        
        for idx, (img_bytes, w, h, ext) in enumerate(batch):
            y_pos = idx * (section_height + vertical_gap)
            rect = fitz.Rect(
                horizontal_margin,
                y_pos + (section_height * 0.05),
                a4_width - horizontal_margin,
                y_pos + section_height - (section_height * 0.05)
            )

            img_aspect = w / h
            rect_aspect = rect.width / rect.height
            
            if img_aspect > rect_aspect:
                new_height = rect.width / img_aspect
                rect.y0 += (rect.height - new_height) / 2
                rect.y1 = rect.y0 + new_height
            else:
                new_width = rect.height * img_aspect
                rect.x0 += (rect.width - new_width) / 2
                rect.x1 = rect.x0 + new_width

            page.insert_image(rect=rect, stream=img_bytes, keep_proportion=False)

    pdf_doc.save(output_path, garbage=4, deflate=True, clean=True)

# Configuration loading
config = load_config()
if config:
    source_dir = config["source_dir"]
    destination_dir = config["destination_dir"]
    extract_images = config["extract_images"]
    convert_pdf = config["convert_pdf"]
    suppress_messages = config["suppress_messages"]
    delete_fdf = config["delete_fdf"]
    coupons_per_page = int(config.get("coupons_per_page", 3))
    merge_pdf = config.get("merge_pdf", False)
else:
    root = get_user_settings()
    root.mainloop()
    config = load_config()
    if not config:
        print("Configuration not set. Exiting script.")
        sys.exit(1)
    source_dir = config["source_dir"]
    destination_dir = config["destination_dir"]
    extract_images = config["extract_images"]
    convert_pdf = config["convert_pdf"]
    suppress_messages = config["suppress_messages"]
    delete_fdf = config["delete_fdf"]
    coupons_per_page = int(config.get("coupons_per_page", 3))
    merge_pdf = config.get("merge_pdf", False)

os.makedirs(destination_dir, exist_ok=True)
operations = {"images_extracted": False, "pdf_created": False}
files_to_delete = []
all_images = []
pdf_files = []

try:
    pdf_files = glob.glob(os.path.join(source_dir, "*.fdf"))
    
    for pdf_file in pdf_files:
        seen_hashes = set()
        unique_images = []
        total_size = 0

        with fitz.open(pdf_file) as doc:
            for page_num in range(doc.page_count):
                page = doc.load_page(page_num)
                img_list = page.get_images(full=True)

                for img_index, img in enumerate(img_list):
                    xref = img[0]
                    base_img = doc.extract_image(xref)
                    img_bytes = base_img["image"]
                    ext = base_img["ext"]
                    w, h = base_img["width"], base_img["height"]

                    if len(img_bytes) < 20480:
                        continue

                    img_hash = get_image_hash(img_bytes)
                    if img_hash in seen_hashes:
                        continue
                    seen_hashes.add(img_hash)

                    processed_bytes = optimize_image_storage(img_bytes, ext)
                    total_size += len(processed_bytes)

                    if extract_images:
                        fname = f"coupon_{os.path.basename(pdf_file)[:-4]}_p{page_num+1}_i{img_index+1}.{ext}"
                        out_path = os.path.join(destination_dir, fname)
                        with open(out_path, "wb") as f:
                            f.write(processed_bytes)
                        operations["images_extracted"] = True

                    unique_images.append((processed_bytes, w, h, ext))

            if convert_pdf and unique_images:
                pdf_path = os.path.join(destination_dir, os.path.basename(pdf_file).replace(".fdf", ".pdf"))
                create_pdf(pdf_path, unique_images, coupons_per_page)
                operations["pdf_created"] = True

            if merge_pdf:
                all_images.extend(unique_images)

        if delete_fdf:
            files_to_delete.append(pdf_file)

    # Only merge if there's more than one FDF file
    if merge_pdf and len(pdf_files) > 1 and all_images:
        merged_path = os.path.join(destination_dir, "merged_coupons.pdf")
        create_pdf(merged_path, all_images, coupons_per_page)
        operations["pdf_created"] = True

    if delete_fdf and files_to_delete:
        for f in files_to_delete:
            try:
                os.remove(f)
            except Exception as e:
                if not suppress_messages:
                    messagebox.showwarning("Deletion Error", f"Error deleting {f}: {str(e)}")

except Exception as e:
    if not suppress_messages:
        messagebox.showerror("Processing Error", f"Critical error: {str(e)}")
    raise

# Show appropriate completion messages
if len(pdf_files) == 0:
    messagebox.showinfo("Complete", "No FDF files found in source directory")
elif not suppress_messages:
    if operations["images_extracted"] or operations["pdf_created"]:
        messagebox.showinfo("Complete", "Extracted successfully!")
    else:
        messagebox.showinfo("Complete", "No processable images found")

print("Operation completed")
