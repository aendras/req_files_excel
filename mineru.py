import os
import json
import torch
from PIL import Image
from transformers import AutoProcessor, Qwen2VLForConditionalGeneration
from mineru_vl_utils import MinerUClient


device = "cuda" if torch.cuda.is_available() else "cpu"
print("Using device:", device)


IMAGE_DIR = r"C:\Users\aendra.shukla\table_vector\output_images\ISO_14229-1_2013.en.PDF"

image_files = sorted([
    os.path.join(IMAGE_DIR, f)
    for f in os.listdir(IMAGE_DIR)
    if f.lower().endswith((".png", ".jpg", ".jpeg", ".tiff", ".bmp"))
])

if not image_files:
    raise RuntimeError("No images found in directory")


model = Qwen2VLForConditionalGeneration.from_pretrained(
    "opendatalab/MinerU2.5-2509-1.2B",
    torch_dtype=torch.float16,
    device_map="auto",
    low_cpu_mem_usage=True
)

processor = AutoProcessor.from_pretrained(
    "opendatalab/MinerU2.5-2509-1.2B",
    use_fast=True
)

client = MinerUClient(
    backend="transformers",
    model=model,
    processor=processor
)


client.generation_kwargs = {
    "max_new_tokens": 256,
    "do_sample": False
}


results = []
MAX_SIDE = 1600  

print("Starting prediction...")

for idx, img_path in enumerate(image_files):
    print(f"\nProcessing [{idx+1}/{len(image_files)}]: {img_path}")

    img = Image.open(img_path).convert("RGB")

   
    if max(img.size) > MAX_SIDE:
        scale = MAX_SIDE / max(img.size)
        img = img.resize(
            (int(img.width * scale), int(img.height * scale)),
            Image.BILINEAR
        )

 
    if img.convert("L").getextrema()[1] < 30:
        print("Skipped (blank page)")
        continue

    with torch.inference_mode():
        blocks = client.two_step_extract(img)

    results.append({
        "image": os.path.basename(img_path),
        "blocks": blocks
    })

    with open("image_extracted_gpu.json", "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)

        
    del blocks
    torch.cuda.empty_cache()





