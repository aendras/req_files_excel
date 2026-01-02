# import os
# import sys
# from huggingface_hub import snapshot_download

# # 1. Automatically find the custom_cert.pem in the same folder
# current_dir = os.path.dirname(os.path.abspath(__file__))
# cert_path = os.path.join(current_dir, "custom_cert.pem")

# # 2. Force Python to use this certificate
# # We set this BEFORE running any network commands
# os.environ["REQUESTS_CA_BUNDLE"] = cert_path
# os.environ["SSL_CERT_FILE"] = cert_path

# print(f"🔒 Using certificate: {cert_path}")
# print("🚀 Starting robust download...")

# try:
#     snapshot_download(
#         repo_id="",
#         local_dir="./Qwen3-14B",
#         local_dir_use_symlinks=False,
#         # Start with 4 workers. If it crashes, change this to 1.
#         max_workers=4,
#         resume_download=True
#     )
#     print("✅ Download complete!")

# except Exception as e:
#     print(f"\n❌ Error: {e}")
#     print("\nIf this failed with an SSL error, your 'custom_cert.pem' might be empty or incorrect.")
#     print("Try changing max_workers=1 in the script.")



import os
import ssl
from huggingface_hub import snapshot_download

# ---------------------------------------------------------
# HARD fallback: disable SSL verification globally
# (Only option when no custom cert exists)
# ---------------------------------------------------------
ssl._create_default_https_context = ssl._create_unverified_context
os.environ["HF_HUB_DISABLE_SSL_VERIFICATION"] = "1"
os.environ["HF_HUB_ENABLE_HF_TRANSFER"] = "1"


print("⚠️ SSL verification DISABLED (no custom cert available)")
print("🚀 Starting model download...")

try:
    snapshot_download(
        repo_id="BAAI/bge-reranker-v2-m3",
        local_dir="./bge-reranker-v2-m3",
        local_dir_use_symlinks=False,
        resume_download=True,
        max_workers=1  # IMPORTANT for corporate networks
    )

    print("✅ Download complete!")

except Exception as e:
    print("\n❌ Download failed")
    print("Reason:", e)
    print("\n💡 If this fails, your network blocks Python HTTPS completely.")
    print("💡 Use browser download and load the model offline.")
