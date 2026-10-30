"""Shared workbook inventory, source extraction and Macro2Chat rebuild jobs."""
from pathlib import Path
import json
import re
import sys
import zipfile
from datetime import datetime

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'vendor' / 'Macro2Chat'))
from macro2chat import OLE_SIGNATURE, OleFile, extract, rebuild


def main():
    request_path = Path(sys.argv[1])
    request = json.loads(request_path.read_text(encoding='utf-8-sig'))
    job = request.pop('job', 'Replace')
    if job == 'Inventory':
        if set(request) != {'paths'} or not isinstance(request['paths'], list) or not all(isinstance(value, str) for value in request['paths']):
            raise ValueError('Invalid workbook inventory request')
        paths = []
        for source in request['paths']:
            path = Path(source)
            with path.open('rb') as handle:
                signature = handle.read(8)
            if signature == OLE_SIGNATURE:
                # Inspect the container only; never decompress VBA dir or module source here.
                ole = OleFile(path.read_bytes())
                vba_ids = {i for i, entry in enumerate(ole.entries) if entry['kind'] == 1 and entry['name'].casefold() == 'vba'}
                has_project = any(entry['kind'] == 2 and entry['name'].casefold() == 'project' for entry in ole.entries)
                has_dir = any(entry['kind'] == 2 and entry['parent'] in vba_ids and entry['name'].casefold() == 'dir' for entry in ole.entries)
                present = has_project and has_dir
            else:
                with zipfile.ZipFile(path) as archive:
                    present = any(entry.filename.replace('\\', '/').rsplit('/', 1)[-1].casefold() == 'vbaproject.bin' for entry in archive.infolist())
            if present:
                paths.append(source)
        print(json.dumps({'paths': paths}, ensure_ascii=False))
        return
    if job not in {'Replace', 'Markdown', 'Modules'} or set(request) != {'source', 'output', 'find', 'replacement', 'match'} or not all(isinstance(value, str) for value in request.values()):
        raise ValueError('Invalid workbook processing request')
    if job == 'Replace' and (not request['find'] or request['match'] not in {'Word', 'Partial'}):
        raise ValueError('A nonempty search and Word/Partial match are required')
    modules = request_path.parent / 'modules'
    extract(request['source'], modules)
    manifest = json.loads((modules / 'manifest.json').read_text(encoding='utf-8'))
    if job == 'Modules':
        result = [dict(name=module['name'], extension=Path(module['file']).suffix, kind=module['kind'], code=(modules / module['file']).read_bytes().decode('utf-8')) for module in manifest['modules']]
        print(json.dumps({'modules': result}, ensure_ascii=False))
        return
    if job == 'Markdown':
        text = ['# ' + Path(request['source']).name + ' — VBA Source Code', '', 'Generated: ' + datetime.now().astimezone().isoformat(), '', '## Module index', '']
        for module in manifest['modules']:
            text.append('- ' + module['name'] + Path(module['file']).suffix + ' (' + module['kind'] + ')')
        for module in manifest['modules']:
            code = (modules / module['file']).read_bytes().decode('utf-8').rstrip('\r\n')
            fence = '`' * max(3, max((len(run) + 1 for run in re.findall('`+', code)), default=3))
            text.extend(['', '## ' + module['name'] + Path(module['file']).suffix, '', fence + 'vb', code, fence])
        with Path(request['output']).open('x', encoding='utf-8-sig', newline='') as output:
            output.write('\r\n'.join(text) + '\r\n')
        print(json.dumps(dict(output=request['output'], changed_modules=len(manifest['modules']), occurrences=0, verified=True), ensure_ascii=False))
        return
    expression = re.escape(request['find'])
    if request['match'] == 'Word':
        expression = r'(?<![A-Za-z0-9_])' + expression + r'(?![A-Za-z0-9_])'
    pattern = re.compile(expression)
    changed_modules = 0
    occurrences = 0
    for module in manifest['modules']:
        path = modules / module['file']
        original = path.read_bytes().decode('utf-8')
        changed, count = pattern.subn(lambda match: request['replacement'], original)
        if count and changed != original:
            changed_modules += 1
            occurrences += count
            path.write_bytes(changed.encode('utf-8'))
    record = dict(output='', changed_modules=changed_modules, occurrences=occurrences, verified=False)
    if occurrences:
        result = rebuild(request['source'], modules, request['output'])
        record.update(output=result['output'], verified=result['verified'], removed_signature_parts=result['removed_signature_parts'])
    print(json.dumps(record, ensure_ascii=False))


if __name__ == '__main__':
    try:
        main()
    except (ValueError, OSError, UnicodeError, KeyError) as error:
        print(json.dumps({'error': str(error)}, ensure_ascii=False), file=sys.stderr)
        raise SystemExit(1)
