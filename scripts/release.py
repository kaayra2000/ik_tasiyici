#!/usr/bin/env python3
import argparse
import sys
import re
import subprocess
from pathlib import Path
import tomllib

def run_command(cmd, check=True):
    print(f"Executing: {' '.join(cmd)}")
    subprocess.run(cmd, check=check)

def main():
    parser = argparse.ArgumentParser(description="Projenin sürüm numarasını arttırır, git commit ve tag işlemlerini yapar.")
    parser.add_argument(
        'bump_type', 
        choices=['major', 'minor', 'patch'], 
        help="Sürüm artış tipi: major (X.0.0), minor (0.X.0) veya patch (0.0.X)"
    )
    parser.add_argument(
        'message', 
        help="Yeni sürüm notu veya değişiklik mesajı"
    )
    args = parser.parse_args()

    bump_type = args.bump_type
    release_notes = args.message
    
    try:
        print("Getting current commit...")
        result = subprocess.run(['git', 'rev-parse', 'HEAD'], capture_output=True, text=True, check=True)
        current_commit = result.stdout.strip()
        print(f"Current commit: {current_commit}")
        
        run_command(['git', 'checkout', 'main'])
        run_command(['git', 'merge', current_commit])
    except subprocess.CalledProcessError:
        print("Git işlemleri sırasında bir hata oluştu (checkout veya merge başarısız).")
        sys.exit(1)
        
    project_root = Path(__file__).parent.parent
    pyproject_path = project_root / 'pyproject.toml'
    
    with open(pyproject_path, 'rb') as f:
        config = tomllib.load(f)
        current_version = config['project']['version']
    
    # Parse version
    match = re.match(r'^(\d+)\.(\d+)\.(\d+)$', current_version)
    if not match:
        print(f"Error: current version '{current_version}' doesn't match x.y.z format")
        sys.exit(1)
        
    major, minor, patch = map(int, match.groups())
    
    if bump_type == 'major':
        major += 1
        minor = 0
        patch = 0
    elif bump_type == 'minor':
        minor += 1
        patch = 0
    elif bump_type == 'patch':
        patch += 1
        
    new_version = f"{major}.{minor}.{patch}"
    print(f"Bumping version: {current_version} -> {new_version}")
    
    # Update pyproject.toml
    content = pyproject_path.read_text(encoding='utf-8')
    content = re.sub(
        r'^version\s*=\s*".*"',
        f'version = "{new_version}"',
        content,
        flags=re.MULTILINE
    )
    pyproject_path.write_text(content, encoding='utf-8')
    
    # Git operations
    try:
        run_command(['git', 'add', 'pyproject.toml'])
        run_command(['git', 'commit', '-m', f"bakım: sürümü v{new_version}'ye yükselt"])
        run_command(['git', 'tag', '-a', f'v{new_version}', '-m', release_notes])
        run_command(['git', 'push', 'origin', 'HEAD'])
        run_command(['git', 'push', 'origin', f'v{new_version}'])
        print(f"Başarıyla yayınlandı: v{new_version}")
    except subprocess.CalledProcessError as e:
        print("Bir hata oluştu.")
        sys.exit(1)

if __name__ == '__main__':
    main()
