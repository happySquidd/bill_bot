from pypdf import PdfReader


key_words = ["due"]


def get_price(single_pdf):

    reader = PdfReader(single_pdf)
    page = reader.pages[0]
    split_text = page.extract_text().split()


    previous_word = ""
    for word in split_text:
        if word[0] == "$" and (previous_word.lower() in key_words) or word[0] == "$" and previous_word.isdigit():
            price = float(word[1:])
        previous_word = word


    return price