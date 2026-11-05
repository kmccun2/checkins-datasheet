# Iterate files in /exports
import os
from openpyxl import load_workbook

# Load mail merge datasheet template
mail_merge_data_wb = load_workbook(filename='utils/Mail Merge Data Template.xlsx')
data_entry_wb = load_workbook(filename='imports/Data Entry.xlsx')
mm_sheet = mail_merge_data_wb.active

# Read data from each sheet in data entry file
cards = []

for sheet in data_entry_wb.worksheets:
    # Read data from each row in the sheet
    for row in sheet.iter_rows(min_row=7, values_only=True):
        if row[0] is not None:  # Make sure data exists

            card_data = {
                'Teacher': sheet.title.split(' - ')[0],
                'Student': row[0],
                'Level': row[1],
                'Score1': row[2],
                'Score2': row[3],
                'Score3': row[4],
                'Score4': row[5],
                'Score5': row[6],
                'Comments': row[7],
                'Tier': sheet.title.split(' - ')[1],
                'Cat1': sheet['C6'].value,
                'Cat2': sheet['D6'].value,
                'Cat3': sheet['E6'].value,
                'Cat4': sheet['F6'].value,
                'Cat5': sheet['G6'].value,
                'Class': sheet['H6'].value,
            }
            cards.append(card_data)

    print("Building list of cards using data from sheet:", sheet.title)

# Order cards by Last Name, First Name (ignoring spacing/case so the same
# student's cards group together even if a sheet was edited by hand) and assign card #
cards = sorted(cards, key=lambda x: (x['Student'].strip().casefold()))
card_no = 1
for card in cards:
    card['Card'] = card_no
    card_no += 1

print("\n","Total cards to process: ", len(cards),"\n")
print("Populating mail merge datasheet...","\n")
for card_data in cards:         
    mm_sheet['A' + str(mm_sheet.max_row + 1)] = card_data['Teacher']  # File name without extension
    mm_sheet['B' + str(mm_sheet.max_row)] = card_data['Student']
    mm_sheet['C' + str(mm_sheet.max_row)] = card_data['Level']
    mm_sheet['D' + str(mm_sheet.max_row)] = card_data['Score1']
    mm_sheet['E' + str(mm_sheet.max_row)] = card_data['Score2']
    mm_sheet['F' + str(mm_sheet.max_row)] = card_data['Score3']
    mm_sheet['G' + str(mm_sheet.max_row)] = card_data['Score4']
    mm_sheet['H' + str(mm_sheet.max_row)] = card_data['Score5']
    mm_sheet['I' + str(mm_sheet.max_row)] = card_data['Comments']
    mm_sheet['J' + str(mm_sheet.max_row)] = card_data['Tier']
    mm_sheet['K' + str(mm_sheet.max_row)] = card_data['Cat1']
    mm_sheet['L' + str(mm_sheet.max_row)] = card_data['Cat2']
    mm_sheet['M' + str(mm_sheet.max_row)] = card_data['Cat3']
    mm_sheet['N' + str(mm_sheet.max_row)] = card_data['Cat4']
    mm_sheet['O' + str(mm_sheet.max_row)] = card_data['Cat5']
    mm_sheet['P' + str(mm_sheet.max_row)] = card_data['Card']


mail_merge_data_wb.save('Mail Merge Datasheet.xlsx')