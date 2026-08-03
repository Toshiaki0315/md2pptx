import os
import argparse
import yaml
import traceback

from generator import PPTXGenerator

def main():
    print("INFO: 変換処理を開始します...")
    parser = argparse.ArgumentParser(description="MarkdownファイルをPowerPointに変換します。")
    parser.add_argument("input", help="変換するMarkdownファイルのパス")
    parser.add_argument("-o", "--output", help="出力ファイル名", default="output.pptx")
    parser.add_argument("-c", "--config", help="YAML設定ファイルのパス", default="config.yaml")
    args = parser.parse_args()
    
    if not os.path.exists(args.input):
        print(f"Error: 入力ファイル '{args.input}' が見つかりません。")
        return

    if not os.path.exists(args.config):
        print(f"Error: 設定ファイル '{args.config}' が見つかりません。")
        return

    try:
        with open(args.config, "r", encoding="utf-8") as f:
            config = yaml.safe_load(f)
            
        # theme設定の適用 (Task 7)
        if 'theme' in config:
            theme = config['theme']
            accent = theme.get('accent_color')
            if accent:
                for k in ['title_h1', 'title_h2', 'title_h3', 'table_header']:
                    if k not in config.get('fonts', {}): config.setdefault('fonts', {})[k] = {}
                    config['fonts'][k]['color_rgb'] = accent
            
            text_color = theme.get('text_color')
            if text_color:
                for k in ['body', 'bullet_level_1', 'table_body']:
                    if k not in config.get('fonts', {}): config.setdefault('fonts', {})[k] = {}
                    config['fonts'][k]['color_rgb'] = text_color
        with open(args.input, "r", encoding="utf-8") as f:
            content = f.read()
            
        generator = PPTXGenerator(config)
        generator.generate(content, args.output)
        
        print(f"Success: '{args.output}' の生成が完了しました！")
        
    except PermissionError:
        print(f"Error: '{args.output}' に書き込めません！PowerPointでファイルを開いたままにしていませんか？閉じてから再実行してください。")
    except Exception as e:
        print(f"Error: 予期せぬエラーが発生しました: {e}")
        traceback.print_exc()

if __name__ == "__main__": 
    main()