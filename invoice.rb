require 'prawn'
require 'yaml'
require 'optparse'
require 'net/smtp'

pwd = File.read('mail.pwd')

options = {}

opt_parser = OptionParser.new do |opt|
  opt.banner = "Usage: invoice.rb [OPTIONS]"
  opt.separator  ""
  opt.separator  "Options"

  opt.on("-s","--setup", "bill for setup and first month") do
    options[:setup] = true
  end

  opt.on("-n","--normal","bill for a normal month") do
    options[:normal] = true
  end

  opt.on("-h","--help","help") do
    puts opt_parser
    exit
  end
end

opt_parser.parse!

config_file = 'normal_month.yml'

if options[:setup]
  config_file = 'setup_month.yml'
end

inv_num = Time.now.strftime('%Y%m%d%H%M')

inv_file = "invoice-#{inv_num}.pdf"

y = YAML.load_file config_file 
total = 0

y.each_with_index do |row, i|
  if i == 0
    next
  end
  
  total = total + row[-1].to_f 
end

Prawn::Document.generate(inv_file) do |pdf|
  
  logopath = 'work_logo.png'
  #initial_y = pdf.cursor
  initialmove_y = 5
  address_x = 35
  invoice_header_x = 325
  lineheight_y = 12
  font_size = 9

  pdf.move_down initialmove_y

  # Add the font style and size
  pdf.font "Helvetica"
  pdf.font_size font_size

  #start with EON Media Group
  pdf.text_box "W.O.R.K.", :at => [address_x,  pdf.cursor]
  pdf.move_down lineheight_y
  pdf.text_box "Kevin Lester", :at => [address_x,  pdf.cursor]
  pdf.move_down lineheight_y
  pdf.text_box "kevin@e-kevin.com", :at => [address_x,  pdf.cursor]
  pdf.move_down lineheight_y
  pdf.text_box "678-357-3319", :at => [address_x,  pdf.cursor]
  pdf.move_down lineheight_y

  last_measured_y = pdf.cursor
  pdf.move_cursor_to pdf.bounds.height

  pdf.image logopath, :width => 215, :position => :right

  pdf.move_cursor_to last_measured_y

  # client address
  pdf.move_down 65
  last_measured_y = pdf.cursor

  pdf.text_box "Finlogic, LLC", :at => [address_x,  pdf.cursor]
  pdf.move_down lineheight_y
  pdf.text_box "Peter Burkes", :at => [address_x,  pdf.cursor]
  pdf.move_down lineheight_y
  pdf.text_box "6030 Bethelview Rd Ste 304", :at => [address_x,  pdf.cursor]
  pdf.move_down lineheight_y
  pdf.text_box "Cumming, GA 30040", :at => [address_x,  pdf.cursor]

  pdf.move_cursor_to last_measured_y

  invoice_header_data = [ 
    ["Invoice #", Time.now.strftime('%Y%m%d') + "000"],
    ["Invoice Date", Time.now.strftime('%A %B %d, %Y')],
    ["Amount Due", '$%.02f' % total]
  ]

  pdf.table(invoice_header_data, :position => invoice_header_x, :width => 215) do
    style(row(0..1).columns(0..1), :padding => [1, 5, 1, 5], :borders => [])
    style(row(2), :background_color => 'e9e9e9', :border_color => 'dddddd', :font_style => :bold)
    style(column(1), :align => :right)
    style(row(2).columns(0), :borders => [:top, :left, :bottom])
    style(row(2).columns(1), :borders => [:top, :right, :bottom])
  end

  pdf.move_down 45

  invoice_services_data = y 

  #format dollar amts
  invoice_services_data.each_with_index do |row, i|
    next if i == 0
    row[-1] = '$%.02f' % row[-1]
  end
  
  #add extra line separator
  invoice_services_data.push ["", "", ""]

  pdf.table(invoice_services_data, :width => pdf.bounds.width) do
    style(row(1..-1).columns(0..-1), :padding => [4, 5, 4, 5], :borders => [:bottom], :border_color => 'dddddd')
    style(row(0), :background_color => 'e9e9e9', :border_color => 'dddddd', :font_style => :bold)
    style(row(0).columns(0..-1), :borders => [:top, :bottom])
    style(row(0).columns(0), :borders => [:top, :left, :bottom])
    style(row(0).columns(-1), :borders => [:top, :right, :bottom])
    style(row(-1), :border_width => 2)
    style(column(2..-1), :align => :right)
    style(columns(0), :width => 75)
    style(columns(1), :width => 375)
  end

  pdf.move_down 1

  invoice_services_totals_data = [ 
    ["Total", '$%.02f' % total],
    ["Amount Paid", "-0.00"],
    ["Amount Due", '$%.02f' % total]
  ]

  pdf.table(invoice_services_totals_data, :position => invoice_header_x, :width => 215) do
    style(row(0..1).columns(0..1), :padding => [1, 5, 1, 5], :borders => [])
    style(row(0), :font_style => :bold)
    style(row(2), :background_color => 'e9e9e9', :border_color => 'dddddd', :font_style => :bold)
    style(column(1), :align => :right)
    style(row(2).columns(0), :borders => [:top, :left, :bottom])
    style(row(2).columns(1), :borders => [:top, :right, :bottom])
  end

  pdf.move_down 25

  invoice_terms_data = [ 
    ["Terms"],
    ["Payable upon receipt"]
  ]

  pdf.table(invoice_terms_data, :width => 275) do
    style(row(0..-1).columns(0..-1), :padding => [1, 0, 1, 0], :borders => [])
    style(row(0).columns(0), :font_style => :bold)
  end

  pdf.move_down 15

  invoice_notes_data = [ 
    ["Notes"],
    ["Thank you for doing business with us."]
  ]

  pdf.table(invoice_notes_data, :width => 275) do
    style(row(0..-1).columns(0..-1), :padding => [1, 0, 1, 0], :borders => [])
    style(row(0).columns(0), :font_style => :bold)
  end

end

marker = "AUNIQUEMARKER"

part1 = <<EOF
From: Kevin Lester <kevin@e-kevin.com>
To: Tracey Lester<traceylester@gmail.com>
Subject: W.O.R.K invoice for #{Time.now.strftime('%B')} 
MIME-Version: 1.0
Content-Type: multipart/mixed; boundary=#{marker}
--#{marker}
EOF

part2 = <<EOF
Content-Type: text/plain
Content-Transfer-Encoding:8bit

Here is your next invoice. Please DO NOT REPLY to this email. If you have any questions please contact 
Kevin Lester
kevin@e-kevin.com
678-357-3319
--#{marker}
EOF

filecontent = File.read(inv_file)
encodedcontent = [filecontent].pack("m")   # base64

part3 = <<EOF
Content-Type: multipart/mixed; name=\”#{inv_file}\”
Content-Transfer-Encoding:base64
Content-Disposition: attachment; filename="#{File.basename inv_file}"
#{encodedcontent}
--#{marker}
EOF
 
message = part1 + part2 + part3
smtp = Net::SMTP.new 'smtp.gmail.com', 587
smtp.enable_starttls
smtp.start('gmail.com', 'work.app.invoice@gmail.com', pwd, :login)
smtp.send_message message, 'kevin@e-kevin.com', 'traceylester@kevin.com', 'kevin.thehick@gmail.com'
smtp.finish
 
