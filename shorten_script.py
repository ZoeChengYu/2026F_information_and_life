import pathlib, re, sys, textwrap

def shorten(text, max_words=10):
    # split on spaces and punctuation, keep Chinese characters as words
    words = re.split(r'\s+', text.strip())
    if len(words) <= max_words:
        return text.strip()
    return ' '.join(words[:max_words])

def process_file(path):
    src = pathlib.Path(path).read_text(encoding='utf-8')
    # pattern for double-quoted string literals (including possible escaped quotes)
    pattern = r'"([^"\\]*?(?:\\.[^"\\]*)*)"'
    def repl(m):
        content = m.group(1)
        # if not long, keep original
        if len(content) <= 30:
            return f'"{content}"'
        short = shorten(content, max_words=8)
        return f'"{short}"'
    new_src = re.sub(pattern, repl, src)
    pathlib.Path(path).write_text(new_src, encoding='utf-8')

if __name__ == '__main__':
    if len(sys.argv) != 2:
        print('Usage: python shorten_script.py <file>')
        sys.exit(1)
    process_file(sys.argv[1])
