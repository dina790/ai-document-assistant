def split_into_chunks(pages, chunk_size=1200, overlap=200):
    """Split PDF pages into smaller sections."""
    chunks = []

    for page_number, text in enumerate(pages, start=1):
        start = 0

        while start < len(text):
            end = min(start + chunk_size, len(text))
            section = text[start:end].strip()

            if section:
                chunks.append({
                    "page": page_number,
                    "text": section,
                })

            if end == len(text):
                break

            start = end - overlap

    return chunks


if __name__ == "__main__":
    sample_pages = [
        "This is the first PDF page.",
        "This is the second PDF page.",
    ]

    result = split_into_chunks(sample_pages)

    for chunk in result:
        print(f"PDF page {chunk['page']}: {chunk['text']}")

    print(f"Created {len(result)} chunks.")