require 'rubygems'
require 'roo'
require 'prawn'

#the file that we'll be loading in - XLSX extension is required
filename_in = 'pauls_numbers.xlsx'
filename_out = 'pauls_book.pdf'

print "Loading '#{filename_in}'... "
if File.readable?(filename_in)
  ss = Excelx.new(filename_in)
  ss.default_sheet = ss.sheets.first
  puts "now in memory"
  puts "Starting PDF creation of #{filename_out}."
  Prawn::Document.generate(filename_out) do |pdf|
    2.upto(ss.last_row) do |line|      
      fname = ss.cell(line,'A').to_s if ss.cell(line,'A')
      lname = ss.cell(line,'B').to_s if ss.cell(line,'B')
      title = ss.cell(line,'C').to_s if ss.cell(line,'C')
      company = ss.cell(line,'D').to_s if ss.cell(line,'D')
      phone1 = ss.cell(line,'E').to_s if ss.cell(line,'E')
      type1 = ss.cell(line,'F').to_s if ss.cell(line,'F')
      phone2 = ss.cell(line,'G').to_s if ss.cell(line,'G')
      type2 = ss.cell(line,'H').to_s if ss.cell(line,'H')
      notes = ss.cell(line,'I').to_s if ss.cell(line,'I')
        
      line1 = nil
      line1 = company unless company == nil     

      if line1 == nil
        line1 = "#{fname} #{lname}"
        if not title == nil and not title.empty?
          line1 = "#{title} #{line1}"
        end
      end
      line2 = "\n"
      line2 = phone1 unless phone1 == nil
      line2 = "#{type1}: #{line2}" unless type1 == nil or type1.empty?

      line3 = "\n"
      line3 = phone2 unless phone2 == nil
      line3 = "#{type2}: #{line3}" unless type2 == nil or type2.empty?

      line4 = "\n"
      line4 = notes unless notes == nil
      if not line1.empty?
        pdf.text("\n#{line1}", :align=>:center, :size=>42)
        pdf.text("#{line2}", :align=>:center, :size=>42)
        pdf.text("#{line3}", :align=>:center, :size=>42)
        pdf.text("#{line4}", :align=>:center, :size=>42)
        
        if line % 2 != 0
          pdf.start_new_page
        else
	  pdf.text("\n", :size=>42)
          pdf.stroke_horizontal_line(25,500)
        end
      end
    end
  end
  puts "Done!"
else
  puts "\nFATAL ERROR: file not readable"
  exit
end
