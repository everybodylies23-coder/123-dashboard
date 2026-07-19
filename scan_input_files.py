import os

dir_path = r"C:\Users\user\Desktop\Antigravity\データ分析自動化"
data_input_dir = os.path.join(dir_path, "data_input")
processed_dir = os.path.join(data_input_dir, "processed")

print("=== Files in data_input ===")
if os.path.exists(data_input_dir):
    for f in os.listdir(data_input_dir):
        p = os.path.join(data_input_dir, f)
        if os.path.isfile(p):
            print(f"  File: {f}, Size: {os.path.getsize(p)}")
            
print("\n=== Files in data_input/processed ===")
if os.path.exists(processed_dir):
    for f in os.listdir(processed_dir):
        p = os.path.join(processed_dir, f)
        if os.path.isfile(p):
            print(f"  File: {f}, Size: {os.path.getsize(p)}")
