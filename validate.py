import xml.etree.ElementTree as ET
import json

# Validate XML
tree = ET.parse('books.xml')
root = tree.getroot()
books = root.findall('book')
print(f'XML valid: root={root.tag}, books={len(books)}')
for b in books:
    print(f"  - {b.find('title').text} by {b.find('author').text}")

# Validate JSON
with open('users.json', 'r', encoding='utf-8') as f:
    users = json.load(f)
print(f'\nJSON valid: users={len(users)}')
for u in users:
    print(f"  - {u['name']}, {u['age']} years old")
