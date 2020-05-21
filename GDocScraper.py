import csv
import gspread
import time

tick = time.perf_counter()
gc = gspread.service_account()

book = gc.open("Copy of OH Debate Tournament - Schedule and Rules - May 2020")

with open('leaguetable.csv', mode='w+', newline="\n") as lt:
    writer = csv.writer(lt, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
    writer.writerow(["Debater", "Game1", "Game2", "Average"])
    lt.flush()
    book_contents = []
    for x in range(3, 14):
        sheet = book.get_worksheet(x)
        sheet_contents = []
        for i in range(9, 36):
            name = "A" + str(i)
            score_g1 = "B" + str(i)
            score_g2_1 = "C" + str(i)
            score_g2_2 = "D" + str(i)
            name_data = sheet.acell(name).value
            score_g1_data = sheet.acell(score_g1).value
            score_g2_1_data = sheet.acell(score_g2_1).value
            score_g2_2_data = sheet.acell(score_g2_2).value

            if name_data and score_g1_data and \
                    'Team' not in name_data and name_data not in range(1, 47):
                if score_g2_1_data and score_g2_2_data:
                    score_g2_data = float(score_g2_1_data) \
                        + float(score_g2_2_data)
                    score_average = (float(score_g1_data)+score_g2_data)/2
                else:
                    score_g2_data, score_average = "", score_g1_data
                print(name_data)
                score_g1_data = float(score_g1_data)
                print(score_g1_data)
                print(score_g2_data)
                sheet_contents.append([name_data, score_g1_data, score_g2_data, score_average])
            time.sleep(3)

        rows = zip(sheet_contents)
        for r in rows:
            writer.writerow(*r)
        lt.flush()
        print(sheet_contents)
        time.sleep(5)

tock = time.perf_counter()
print(tock - tick)
