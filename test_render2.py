import pandas as pd
import os
import asyncio
import base64
from jinja2 import Environment, FileSystemLoader
from playwright.async_api import async_playwright
import sys

# Import from generate2
sys.path.insert(0, r'C:\Users\chlgn\OneDrive\Desktop\general prize')
from generate2 import load_data, build_context

BASE_DIR   = r'C:\Users\chlgn\OneDrive\Desktop\general prize'
OUTPUT_DIR = os.path.join(BASE_DIR, 'output_test_2')
TEMPLATE   = 'template2.html'

async def test_render():
    print("Loading data...")
    df_f, s_map, b_map = load_data()
    
    # Process top 10 rows for testing
    df_head = df_f.head(10)
    
    env  = Environment(loader=FileSystemLoader(BASE_DIR))
    tmpl = env.get_template(TEMPLATE)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    # CI Load
    ci_path = os.path.join(BASE_DIR, 'ci.png')
    ci_b64 = ''
    if os.path.exists(ci_path):
        with open(ci_path, 'rb') as f:
            ci_b64 = 'data:image/png;base64,' + base64.b64encode(f.read()).decode()

    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page    = await browser.new_page(viewport={'width': 1600, 'height': 1100})
        
        for idx, (_, row) in enumerate(df_head.iterrows(), 1):
            ctx = build_context(row, s_map, b_map) # Pass maps
            
            html_str = tmpl.render(**ctx, ci_src=ci_b64)
            await page.set_content(html_str, wait_until='networkidle')
            
            out = os.path.join(OUTPUT_DIR, f'test_{idx}.png')
            await page.screenshot(path=out, clip={'x':0,'y':0,'width':1600,'height':1100})
            print(f'[{idx}/3] {out}')
            
        await browser.close()
    print('Test done.')

if __name__ == '__main__':
    asyncio.run(test_render())
