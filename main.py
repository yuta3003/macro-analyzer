from oletools.olevba import VBA_Parser

vba_parser = VBA_Parser("計測器貸出管理.xlsm")
if vba_parser.detect_vba_macros():
    for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
        print(f"--- {vba_filename} ---")
        print(vba_code)
