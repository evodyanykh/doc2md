from pypdf import PdfReader, PdfWriter

reader = PdfReader("ranepa1/Использование БРС.pdf")
writer = PdfWriter()

for i in range(0, 2):  # страницы 20–32
    writer.add_page(reader.pages[i])

with open("1-3.pdf", "wb") as f:
    writer.write(f)

print("Готово")
