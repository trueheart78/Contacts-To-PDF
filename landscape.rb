require 'rubygems'
require 'roo'
require 'prawn'

#the file that we'll be loading in - XLSX extension is required
filename_in = 'pauls_numbers.xlsx'
filename_out = 'pauls_book-landscape.pdf'

font_main = 72
font_size = 64

print "Loading '#{filename_in}'... "
if File.readable?(filename_in)
  ss = Excelx.new(filename_in)
  ss.default_sheet = ss.sheets.first
  puts "now in memory"
  puts "Starting PDF creation of #{filename_out}."
  Prawn::Document.generate(filename_out, :page_layout=>:landscape) do |pdf|
    
    2.upto(ss.last_row) do |line|      
      fname = ss.cell(line,'A').to_s.upcase.rstrip if ss.cell(line,'A')
      lname = ss.cell(line,'B').to_s.upcase.lstrip if ss.cell(line,'B')
      title = ss.cell(line,'C').to_s.upcase.rstrip if ss.cell(line,'C')
      company = ss.cell(line,'D').to_s.upcase if ss.cell(line,'D')
      phone1 = ss.cell(line,'E').to_s.upcase if ss.cell(line,'E')
      type1 = ss.cell(line,'F').to_s.upcase if ss.cell(line,'F')
      phone2 = ss.cell(line,'G').to_s.upcase if ss.cell(line,'G')
      type2 = ss.cell(line,'H').to_s.upcase if ss.cell(line,'H')
      id = ss.cell(line,'I').to_s.upcase if ss.cell(line,'I')
      notes = ss.cell(line,'J').to_s.upcase if ss.cell(line,'J')        

      line1 = nil
      line1 = company unless company == nil     

      if line1 == nil
        line1 = "#{fname} #{lname}"
        if not title == nil and not title.empty?
          line1 = "#{title} #{line1}"
        end
      elsif not fname == nil and not lname == nil and not fname.empty? and not lname.empty?
        if not title == nil and not title.empty?
          line1 = "#{line1}\n#{title} #{fname} #{lname}"
        else
          line1 = "#{line1}\n#{fname} #{lname}"
        end
      end

      line2 = "\n"
      line2 = phone1 unless phone1 == nil
      
      line3 = "\n"
      line3 = "(#{type1})" unless type1 == nil or type1.empty?

      line4 = "\n"
      line4 = phone2 unless phone2 == nil
      
      line5 = "\n"
      line5 = "(#{type2})" unless type2 == nil or type2.empty?

      line6 = "\n"
      line6 = "\nID: #{id}" unless id == nil

      line7 = "\n"
      line7 = notes unless notes == nil
      if not line1.empty?
        pdf.font "Courier"
        pdf.text("#{line1}", :align=>:center, :size=>font_main, :style=>:bold)
        pdf.text("#{line2}", :align=>:center, :size=>font_main, :style=>:bold)
        pdf.text("#{line3}", :align=>:center, :size=>font_main, :style=>:bold)
        pdf.text("#{line4}", :align=>:center, :size=>font_size, :style=>:bold)
        pdf.text("#{line5}", :align=>:center, :size=>font_size, :style=>:bold)
        pdf.text("#{line6}", :align=>:center, :size=>font_size, :style=>:bold)
        pdf.text("#{line7}", :align=>:center, :size=>font_size, :style=>:bold)
        pdf.start_new_page
      end
    end
  end
  puts "Done!"
else
  puts "\nFATAL ERROR: file not readable"
  exit
end
