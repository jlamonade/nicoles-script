import os
import xlsxwriter

workbook = xlsxwriter.Workbook('../hello_world.xlsx') # change output filename here
worksheet = workbook.add_worksheet()
template_sheet = ( # column names can be changed
  ["ID"], ["TIME PERIOD"], ["SCAN"], ["TIME START"], ["TIME END"]
)

def parse_file(): 
  # parses .dat files in specific format within the ./files directory
  # folder must contain only .dat files
  # directory can be changed from "files"

  for filename in os.listdir(os.chdir("files")):
    with open(os.path.join(os.getcwd(), filename), 'r') as f:

      lines = [line for line in f]

      sample_name = ' '.join(lines[34].split()[2:4])
      time_period = lines[34].split()[4]
      scan = lines[34].split()[5]
      time_start = ' '.join(lines[41].split()[2:4])
      time_end = ' '.join(lines[42].split()[2:4])

      template_sheet[0].append(sample_name)
      template_sheet[1].append(time_period)
      template_sheet[2].append(scan)
      template_sheet[3].append(time_start)
      template_sheet[4].append(time_end)


def create_sheet():
  # writes data to the worksheet

  col = 0

  for column in template_sheet:
    row = 0
    worksheet.write(row, col, column[0])
    for value in column[1:]:
      worksheet.write(row + 1, col, value)
      row += 1
    col += 1

  workbook.close()

def main():
  # main function

  parse_file()
  create_sheet()

main()