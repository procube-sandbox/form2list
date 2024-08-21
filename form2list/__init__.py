#!/usr/bin/env python

import argparse
import os
import openpyxl
import yaml
from jinja2 import Template

def parse_arguments():
    parser = argparse.ArgumentParser(description='ディレクトリを巡回して入力ファイルを探します。')
    parser.add_argument('-c', '--config', type=str, default='spec.yml', help='変換仕様定義YAMLファイル名')
    parser.add_argument('-o', '--output', type=str, default='list.xlsx', help='出力ファイル')
    parser.add_argument('-v', '--verbose', action='store_true', help='詳細な出力を表示')
    parser.add_argument('directory', type=str, help='読み込みフォルダのパス')
    return parser.parse_args()

def find_input_files(directory):
    input_files = []
    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith('.xlsx'):  # 入力ファイルの拡張子を指定
                input_files.append(os.path.join(root, file))
    return input_files

def verbose_print(verbose, message):
    if verbose:
        print(message)

def process_file(file_path, config, output_wb, rawNumber, verbose):
    verbose_print(verbose, f'Processing file: {file_path}')
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    context = {}
    found = False
    for spec in config['inputFormats']:
        for key, cell in spec['items'].items():
            context[key] = ws[cell].value
        verbose_print(verbose, f"Check condition {spec.get('condition', 'True')} on context {context}")
        if spec['condition_template'].render(context):
            found = True
            break
    if not found:
        print(f"Fail to load values from: {file_path}", file=sys.stderr)
        return False
    context['dirname'] = os.path.dirname(file_path)
    context['basename'] = os.path.basename(file_path)

    for sheet_config in config['sheets']:
        ws = output_wb[sheet_config.name]
        columnNumber = 1
        for col in sheet_config.columns:
            cell = ws.cell(cell=columnNumber + config.get('columnOffset', 0), raw=rawNumber + config.get('rawOffset', 0))
            cell.value = col["value_template"].render(context)
    return True

def setup_templates(config):
    for spec in config['inputFormats']:
        condition = spec.get('condition', 'True')
        spec['condition_template'] = Template(condition)
    for sheet_config in config['sheets']:
        for col in sheet_config.get('columns', []):
            col['value_template'] = Template(col['value'])

def main():
    args = parse_arguments()
    input_files = find_input_files(args.directory)
    
    if not input_files:
        print('入力ファイルが見つかりませんでした。', file=sys.stderr)
        return  1
    
    with open(args.config, 'r') as yaml_file:
        config = yaml.safe_load(yaml_file)
    setup_templates(config)

    output_wb = openpyxl.load_workbook(args.output)
    raw_number = 1
    for file_path in input_files:
        result = process_file(file_path, config, output_wb, raw_number, args.verbose)
        raw_number += 1
        if not result:
            return 1
    output_wb.save(args.output)
    return 0

if __name__ == '__main__':
    main()