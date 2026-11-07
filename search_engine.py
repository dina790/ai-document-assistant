import re
from document_search import split_into_chunks


def find_relevant_chunks(chunks, question, limit=4):
    """Return sections sharing keywords with the question."""
    stop_words = {
        "a", "an", "the", "is", "are", "was", "were",
        "what", "which", "who", "how", "why", "where",
        "in", "on", "of", "to", "for", "and", "or",
        "this", "that", "it", "does", "do", "did",
        "pdf", "document", "please",
    }

    question_words = set(re.findall(r"\w+", question.lower()))
    keywords = question_words - stop_words

    matches = []

    for chunk in chunks:
        chunk_words = set(re.findall(r"\w+", chunk["text"].lower()))
        score = len(keywords & chunk_words)

        if score > 0:
            matches.append((score, chunk))

    matches.sort(key=lambda item: item[0], reverse=True)

    return [chunk for score, chunk in matches[:limit]]


if __name__ == "__main__":
    pages = [
        "Python is a programming language.",
        "Solar panels generate electricity from sunlight.",
        "A database stores information.",
    ]

    chunks = split_into_chunks(pages)
    results = find_relevant_chunks(
        chunks,
        "How do solar panels generate electricity?",
    )

    for result in results:
        print(f"PDF page {result['page']}: {result['text']}")

    print(f"Found {len(results)} matching section(s).")