require 'rubygems'
require 'parseexcel'
require 'fastercsv'

caja_chica_path = "caja_chica.xls"

# open file
wb = Spreadsheet::ParseExcel.parse(caja_chica_path)

FasterCSV.open("caja_chica.csv","w") do |csv|
  csv << %W(date teacher desc inc out concept)

  # iterate through sheets
  (2...wb.sheet_count).each do |sn|
    ws = wb.worksheet(sn)
    date = ws.row(0).at(1).date
    # iterate through data box in work sheet
    (4...74).each do |n|
      row = ws.row(n)

      teacher = row.at(3).to_s # teacher
      desc    = row.at(4).to_s # description
      inc     = row.at(5).to_f # income
      out     = row.at(6).to_f # outcome
      concept = row.at(7).to_s # concept

      # ignore lines without income or outcome
      unless inc==0.0 && out==0.0
        csv << [date.to_s,teacher.to_s,desc.to_s,inc.to_s,out.to_s,concept.to_s]
      end
    end
  end
end
