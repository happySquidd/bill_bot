from pypdf import PdfReader


months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def get_dates(single_pdf):

    reader = PdfReader(single_pdf)
    page = reader.pages[0]
    split_text = page.extract_text().split()

    # print(page.extract_text())
    bill_date = ""
    from_date = ""
    to_date = ""
    idx = 0
    for word in split_text:
        if word in months:
            day = str(split_text[idx+1]).replace(",", "")

            for i in day:
                try:
                    int(i)
                except ValueError:
                    day = day.replace(i, "")

            date = word + " " + day
            if bill_date == "":
                bill_date = date
            elif from_date == "":
                from_date = date
            elif to_date == "" and date != from_date:
                to_date = date
        idx += 1
    
    return bill_date, from_date, to_date
