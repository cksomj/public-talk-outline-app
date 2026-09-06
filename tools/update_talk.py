#!/usr/bin/env python3
"""
update_talk.py — 새/수정된 골자 .docx 파일을 앱(index.html)의 TALKS 데이터에 자동 반영한다.

사용법은 이 폴더의 README.md 참조. 요약:
    python update_talk.py <파일1.docx> [파일2.docx ...]
    python update_talk.py --dir             (기본 new_talks 폴더 안의 모든 .docx 처리)

GOLJA-AUTO-UPDATE-TOOL-20260906-G7에서 작성됨. G1~G3에서 검증된 파싱 규칙
(process.py / process_g3.py)을 그대로 이식했다 — 새로 추측한 규칙 없음.
"""
import argparse
import glob
import json
import os
import re
import subprocess
import sys
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime

# Windows 콘솔 기본 코드페이지(cp949)는 em-dash(—) 등 일부 유니코드 문자를 못
# 담아서 --help/print 출력이 UnicodeEncodeError로 죽는 문제가 있어, 시작 시
# stdout/stderr을 UTF-8로 강제 전환한다(GOLJA-COMMIT-PUSH-ALL-20260906-G8에서
# --help 실행 중 실제로 재현/발견함).
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, 'reconfigure'):
        try:
            _stream.reconfigure(encoding='utf-8')
        except Exception:
            pass

# 이 파일은 저장소 안(<repo>/tools/update_talk.py)에 들어 있으므로, 저장소 루트는
# 한 단계 위다(GOLJA-COMMIT-PUSH-ALL-20260906-G8에서 저장소 안으로 옮겨 넣음).
REPO_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INDEX_PATH = os.path.join(REPO_DIR, 'index.html')
# 체크포인트/새 골자 입력 폴더는 저장소 밖(작업 폴더)에 둔다 - 원본 .docx나 백업
# 파일을 공개 저장소에 커밋하지 않기 위함. 이 폴더들이 없으면 자동으로 만든다.
WORKROOT = os.path.dirname(REPO_DIR)
CHECKPOINT_DIR = os.path.join(WORKROOT, 'checkpoints')
DEFAULT_NEW_TALKS_DIR = os.path.join(WORKROOT, 'new_talks')

# 앱의 TH 배열(index.html) 순서와 반드시 같아야 한다. 신규 추가된 골자의 테마
# 분류는 원본 .docx만 봐서는 알 수 없어서(편집자 판단 영역), 기본값을 넣어 두되
# 반드시 사람이 다시 확인하라고 크게 경고한다 - t가 null이면 앱이 그 카드를 열 때
# 바로 오류가 나는 것을 실제로 확인했다(TH[tk.t]가 undefined가 됨).
DEFAULT_THEME_INDEX = 0
THEME_NAMES = ['하느님·그리스도', '마지막 날·예언', '가정·자녀', '도덕·행실', '성경·진리·종교',
               '기도·영성', '구원·부활', '전도·봉사', '신뢰·믿음·인내', '사회·윤리']

W_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

ZWSP_RE = re.compile('[\u200B\uFEFF\u200C\u200D]')
HEADING_RE = re.compile(r'\s*\(\d+\s*분\)\s*$')
HEADING_RE_LOOSE = re.compile(r'\s*\d+\s*분\s*$')
NOTE_LABEL_RE = re.compile(r'^(?:연사의\s*)?유의\s*사항\s*[:：]?\s*(.*)$')
TITLE_PREFIX_RE = re.compile(r'^제\s*\d+\s*번\s*')
FILENAME_NO_RE = re.compile(r'S-3?4?[_-]?KO[_-](\d+)', re.IGNORECASE)
FILENAME_NO_FALLBACK_RE = re.compile(r'(\d+)(?=\.[^.]+$)')


def clean_text(s):
    return ZWSP_RE.sub('', s).replace('\r', '').strip()


def extract_docx_paragraphs(path):
    """docx의 word/document.xml에서 문단(w:p)마다 w:t 텍스트만 이어 붙인다.
    w:br(줄바꿈)/w:tab은 의도적으로 무시한다 - G2에서 실제 docx XML을 열어
    확인한 결과, 이 자리는 원문에서도 공백 없이 그대로 이어지는 것이 맞았다."""
    with zipfile.ZipFile(path) as z:
        with z.open('word/document.xml') as f:
            xml_bytes = f.read()
    root = ET.fromstring(xml_bytes)
    body = root.find(f'{W_NS}body')
    paras = []
    for p in body.findall(f'{W_NS}p'):
        texts = [t.text or '' for t in p.iter(f'{W_NS}t')]
        line = clean_text(''.join(texts))
        if line:
            paras.append(line)
    return paras


def build_sections(paras, idx, heading_re):
    sections = []
    cur = None
    for para in paras[idx:]:
        if heading_re.search(para):
            cur = {'title': heading_re.sub('', para).strip(), 'points': []}
            sections.append(cur)
        elif cur is not None:
            cur['points'].append(para)
        else:
            cur = {'title': '(unknown-lead-in)', 'points': [para]}
            sections.append(cur)
    return sections


def parse_docx(path):
    paras = extract_docx_paragraphs(path)
    if not paras:
        raise ValueError(f'{path}: 문단을 하나도 추출하지 못했습니다(빈 문서이거나 형식이 다름)')
    title = TITLE_PREFIX_RE.sub('', paras[0]).strip()
    idx = 1
    note = ''
    m = NOTE_LABEL_RE.match(paras[1]) if len(paras) > 1 else None
    if m:
        if m.group(1) and m.group(1).strip():
            note = m.group(1).strip()
            idx = 2
        else:
            note = paras[2] if len(paras) > 2 else ''
            idx = 3

    sections = build_sections(paras, idx, HEADING_RE)
    used_loose = False
    if len(sections) <= 1 and sections and sections[0]['title'] == '(unknown-lead-in)':
        loose_sections = build_sections(paras, idx, HEADING_RE_LOOSE)
        if len(loose_sections) > 1:
            sections = loose_sections
            used_loose = True

    return {'title': title, 'note': note, 'sections': sections}, used_loose


def extract_no_from_filename(path):
    base = os.path.basename(path)
    m = FILENAME_NO_RE.search(base)
    if m:
        return int(m.group(1))
    m = FILENAME_NO_FALLBACK_RE.search(base)
    if m:
        return int(m.group(1))
    return None


def load_index_html():
    with open(INDEX_PATH, encoding='utf-8', newline='') as f:
        return f.read()


def find_data_block(html):
    m = re.search(r'(<script id="DATA" type="application/json">)([\s\S]*?)(</script>)', html)
    if not m:
        raise RuntimeError('index.html에서 DATA 스크립트 블록을 찾지 못했습니다')
    return m


def make_checkpoint():
    os.makedirs(CHECKPOINT_DIR, exist_ok=True)
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    dst = os.path.join(CHECKPOINT_DIR, f'index.html.before_update_talk_{ts}.bak')
    with open(INDEX_PATH, 'rb') as fsrc, open(dst, 'wb') as fdst:
        fdst.write(fsrc.read())
    return dst


def upsert_talk(talks, no, parsed, theme_index=None):
    """no가 이미 있으면 그 자리에서 교체한다 - title/note/sections만 갱신하고 t(테마
    분류)는 절대 건드리지 않는다(중요: 한 번의 실행에 업데이트 대상과 신규 추가
    대상이 섞여 있을 때 --theme이 신규 추가용으로 지정된 것이 무관한 기존 항목의
    테마까지 덮어써 버리는 사고가 실제로 나서 이렇게 분리함 - 기존 항목의 테마를
    고치고 싶으면 --set-theme NO:N 을 따로 써야 한다). no가 없으면 새로 삽입하고,
    이때만 theme_index(없으면 기본값)를 적용한다 - t가 null이면 앱이 그 카드를 열 때
    즉시 오류가 나는 것을 실제로 확인했기 때문에 null로 두지 않는다."""
    for i, tk in enumerate(talks):
        if tk['no'] == no:
            tk['title'] = parsed['title']
            tk['note'] = parsed['note']
            tk['sections'] = parsed['sections']
            return 'updated', i, False
    assigned_theme = theme_index if theme_index is not None else DEFAULT_THEME_INDEX
    new_talk = {'no': no, 'title': parsed['title'], 't': assigned_theme, 'note': parsed['note'], 'sections': parsed['sections']}
    insert_at = len(talks)
    for i, tk in enumerate(talks):
        if tk['no'] > no:
            insert_at = i
            break
    talks.insert(insert_at, new_talk)
    return 'inserted', insert_at, (theme_index is None)


def write_talks(html, m, talks):
    leading_ws = re.match(r'^\s*', m.group(2)).group(0)
    trailing_ws = re.search(r'\s*$', m.group(2)).group(0)
    new_json = json.dumps(talks, ensure_ascii=False, separators=(', ', ': '))
    new_html = html[:m.start()] + m.group(1) + leading_ws + new_json + trailing_ws + m.group(3) + html[m.end():]
    return new_html


def validate(html):
    m = find_data_block(html)
    talks = json.loads(m.group(2))
    if not isinstance(talks, list):
        raise ValueError('TALKS가 배열이 아닙니다')
    nos = [t['no'] for t in talks]
    if len(nos) != len(set(nos)):
        raise ValueError('중복된 no가 있습니다: ' + str([n for n in nos if nos.count(n) > 1]))
    if nos != sorted(nos):
        raise ValueError('no 순서가 정렬되어 있지 않습니다')
    return talks


def take_screenshot(no, out_dir):
    """로컬 서버(http.server)를 잠깐 띄우고 헤드리스 Chrome으로 해당 번호 상세 화면을
    캡처한다. Chrome/Node가 없으면 건너뛰고 경고만 남긴다(스크립트 자체는 계속 진행)."""
    node = _find_exe('node')
    chrome = _find_chrome()
    if not node or not chrome:
        print(f'  [경고] node 또는 Chrome을 찾지 못해 talk {no} 스크린샷을 건너뜁니다')
        return None

    port = 5299
    server = subprocess.Popen(
        [sys.executable, '-m', 'http.server', str(port), '--bind', '127.0.0.1'],
        cwd=REPO_DIR, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )
    try:
        import time
        import urllib.request
        for _ in range(20):
            try:
                urllib.request.urlopen(f'http://127.0.0.1:{port}/', timeout=1)
                break
            except Exception:
                time.sleep(0.3)

        driver_js = os.path.join(os.path.dirname(os.path.abspath(__file__)), '_shot_driver.js')
        out_path = os.path.join(out_dir, f'talk{no}.png')
        result = subprocess.run(
            [node, driver_js, f'http://127.0.0.1:{port}/', f'shP({no})', out_path],
            capture_output=True, text=True, timeout=30,
        )
        if result.returncode != 0:
            print(f'  [경고] talk {no} 스크린샷 실패: {result.stderr.strip()[:200]}')
            return None
        return out_path
    finally:
        server.terminate()
        server.wait(timeout=5)


def _find_exe(name):
    from shutil import which
    return which(name)


def _find_chrome():
    candidates = [
        r'C:\Program Files\Google\Chrome\Application\chrome.exe',
        r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def git_run(args, cwd=REPO_DIR):
    result = subprocess.run(['git'] + args, cwd=cwd, capture_output=True, text=True)
    return result.returncode, result.stdout.strip(), result.stderr.strip()


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('files', nargs='*', help='.docx 파일 경로(들)')
    ap.add_argument('--dir', nargs='?', const=DEFAULT_NEW_TALKS_DIR,
                     help=f'이 폴더 안의 모든 .docx를 처리. 값을 안 주면 기본 폴더'
                          f'({DEFAULT_NEW_TALKS_DIR})를 사용')
    ap.add_argument('--no-commit', action='store_true', help='git commit을 하지 않음(반영만)')
    ap.add_argument('--no-push', action='store_true', help='git commit은 하되 push는 하지 않음')
    ap.add_argument('--no-screenshot', action='store_true', help='스크린샷 검증 생략')
    ap.add_argument('--theme', type=int, choices=range(10), metavar='0-9',
                     help='이번 실행에서 "신규 추가"되는 항목에만 적용할 테마 분류 번호'
                          '(0=하느님·그리스도 ... 9=사회·윤리). 지정하지 않으면 0으로 넣고 '
                          '검토가 필요하다고 표시함. 이미 있는 번호(업데이트 대상)에는 '
                          '절대 영향을 주지 않음')
    ap.add_argument('--set-theme', action='append', default=[], metavar='NO:테마번호',
                     help='이미 있는 번호의 테마를 나중에 바로잡을 때만 쓴다(예: --set-theme 195:3). '
                          '여러 번 반복 가능. .docx 처리와는 완전히 별개로 동작함')
    args = ap.parse_args()

    files = list(args.files)
    if args.dir:
        files += sorted(glob.glob(os.path.join(args.dir, '*.docx')))
    files = [f for f in files if not os.path.basename(f).startswith('~$')]  # 워드 임시파일 제외
    if not files and not args.set_theme:
        print('처리할 .docx 파일이 없습니다. 파일 경로를 인자로 주거나 --dir 옵션을 쓰세요.')
        sys.exit(1)

    set_theme_pairs = []
    for spec in args.set_theme:
        try:
            no_s, theme_s = spec.split(':')
            set_theme_pairs.append((int(no_s), int(theme_s)))
        except Exception:
            print(f'[오류] --set-theme 형식이 잘못됨: {spec!r} (예: 195:3)')
            sys.exit(1)
        if not (0 <= set_theme_pairs[-1][1] <= 9):
            print(f'[오류] --set-theme의 테마 번호는 0~9여야 합니다: {spec!r}')
            sys.exit(1)

    if files:
        print(f'== 대상 파일 {len(files)}개 ==')
        for f in files:
            print(' -', f)

    checkpoint_path = make_checkpoint()
    print(f'\n체크포인트 생성: {checkpoint_path}')

    html = load_index_html()
    m = find_data_block(html)
    talks = json.loads(m.group(2))
    before_count = len(talks)

    results = []
    for path in files:
        no = extract_no_from_filename(path)
        if no is None:
            print(f'[건너뜀] {path}: 파일명에서 번호를 추출하지 못했습니다')
            continue
        try:
            parsed, used_loose = parse_docx(path)
        except Exception as e:
            print(f'[오류] {path} (no={no}): 파싱 실패 - {e}')
            continue
        action, pos, needs_theme_review = upsert_talk(talks, no, parsed, args.theme)
        results.append({'no': no, 'action': action, 'path': path, 'used_loose': used_loose,
                         'needs_theme_review': needs_theme_review,
                         'section_count': len(parsed['sections']),
                         'point_count': sum(len(s['points']) for s in parsed['sections'])})
        loose_note = ' (공백 있는 "(N분)" 보정 사용)' if used_loose else ''
        print(f'  no={no}: {action}{loose_note} - 섹션 {len(parsed["sections"])}개, '
              f'요점 {sum(len(s["points"]) for s in parsed["sections"])}개')

    theme_changes = []
    for no, theme_n in set_theme_pairs:
        tk = next((t for t in talks if t['no'] == no), None)
        if tk is None:
            print(f'[오류] --set-theme: no={no} 항목이 없습니다')
            sys.exit(1)
        old_t = tk.get('t')
        tk['t'] = theme_n
        theme_changes.append(no)
        print(f'  no={no}: 테마 {old_t} -> {theme_n} ({THEME_NAMES[theme_n]})')

    if not results and not theme_changes:
        print('\n반영된 항목이 없습니다. 종료합니다.')
        sys.exit(1)

    theme_review_needed = [r['no'] for r in results if r['needs_theme_review']]
    if theme_review_needed:
        print('\n' + '=' * 60)
        print(f'[꼭 확인] 신규 추가된 번호 {theme_review_needed}는 테마 분류를 지정하지')
        print(f'않아서 기본값(0: {THEME_NAMES[DEFAULT_THEME_INDEX]})으로 넣었습니다.')
        print('앱의 "편집" 화면에는 테마를 바꾸는 입력란이 없으므로(직접 확인함),')
        print('실제 주제에 맞게 바로잡으려면 이 도구를 아래처럼 다시 실행하세요:')
        for no in theme_review_needed:
            print(f'    python update_talk.py --set-theme {no}:N')
        print('(테마 없이 null로 두면 앱에서 그 카드를 열 때 바로 오류가 납니다.)')
        print('테마 번호: ' + ', '.join(f'{i}={n}' for i, n in enumerate(THEME_NAMES)))
        print('=' * 60)

    new_html = write_talks(html, m, talks)

    try:
        validated_talks = validate(new_html)
    except Exception as e:
        print(f'\n[치명적 오류] 반영 후 JSON 검증 실패: {e}')
        print('index.html은 변경하지 않았습니다.')
        sys.exit(1)

    with open(INDEX_PATH, 'w', encoding='utf-8', newline='') as f:
        f.write(new_html)

    print(f'\nindex.html 갱신 완료. talks 개수: {before_count} -> {len(validated_talks)}')

    if not args.no_screenshot:
        print('\n== 스크린샷 검증 ==')
        shot_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'verify_screenshots')
        os.makedirs(shot_dir, exist_ok=True)
        shot_targets = sorted(set([r['no'] for r in results] + theme_changes))
        for no in shot_targets:
            shot = take_screenshot(no, shot_dir)
            if shot:
                print(f'  talk {no}: {shot}')

    updated = [r['no'] for r in results if r['action'] == 'updated']
    inserted = [r['no'] for r in results if r['action'] == 'inserted']
    msg_lines = ['골자 자동 반영 (update_talk.py)']
    if updated:
        msg_lines.append('- 교체된 번호: ' + ', '.join(str(n) for n in updated))
    if inserted:
        msg_lines.append('- 신규 추가된 번호: ' + ', '.join(str(n) for n in inserted))
    if theme_changes:
        msg_lines.append('- 테마 재지정된 번호: ' + ', '.join(str(n) for n in theme_changes))
    commit_msg = '\n'.join(msg_lines)

    if args.no_commit:
        print('\n--no-commit 지정됨: git 작업 생략')
        print('커밋했다면 사용했을 메시지:\n' + commit_msg)
        return

    print('\n== git 작업 ==')
    rc, out, err = git_run(['add', 'index.html'])
    if rc != 0:
        print('[오류] git add 실패:', err)
        sys.exit(1)
    rc, out, err = git_run(['commit', '-m', commit_msg])
    if rc != 0:
        print('[오류] git commit 실패:', err)
        sys.exit(1)
    print('커밋 완료:\n' + out)

    if args.no_push:
        print('\n--no-push 지정됨: 커밋만 하고 push는 하지 않았습니다.')
        return

    rc, out, err = git_run(['remote', '-v'])
    if 'origin' not in out:
        print('\ngit remote(origin)가 연결되어 있지 않아 push를 생략합니다.')
        return

    rc, out, err = git_run(['push', 'origin', 'HEAD'])
    if rc != 0:
        print('[오류] git push 실패:', err)
        print('로컬 커밋은 완료된 상태입니다. 수동으로 push해 주세요.')
        sys.exit(1)
    print('push 완료:\n' + out)


if __name__ == '__main__':
    main()
