import re

def allnovelbook_extract_title(text):
    # Ищем название после "You’re Reading" и до "Novel"
    print(f'[[{text}]]')
    match = re.search(r"Вы читаете\s+(.*?)\s+на AllNovelBook.com", text)
    if match:
        res = match.group(1).strip()
        res = res.replace('роман', '')
        res = res.replace('«', '')
        res = res.replace('»', '')
        res = res.strip()
        return res
    return None