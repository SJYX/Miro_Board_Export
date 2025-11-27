import os

file_path = "miro_export_report.csv"

if os.path.exists(file_path):
    try:
        with open(file_path, 'rb') as f:
            content = f.read()
            print(f"File size: {len(content)} bytes")
            print(f"First 100 bytes: {content[:100]}")
            try:
                decoded = content.decode('utf-8-sig')
                print(f"Decoded first 100 chars: {decoded[:100]}")
            except Exception as e:
                print(f"UTF-8-SIG decode failed: {e}")
                try:
                    decoded = content.decode('utf-16')
                    print(f"UTF-16 decode success: {decoded[:100]}")
                except:
                    print("UTF-16 decode failed")
    except Exception as e:
        print(f"Read failed: {e}")
else:
    print("File not found")
