# frozen_string_literal: true

require 'optparse'
require 'ooxml_parser'

# Parse command line option
def parse_command_line
  options = {}

  OptionParser.new do |opts|
    opts.on('--file FILE', 'Specify the DOCX file') do |file|
      options[:file] = file
    end
  end.parse!

  if options[:file].nil?
    puts 'Error: Please provide a DOCX file using --file parameter.'
    exit(255)
  end
  options
end

# Verify the file
def verification(file_path)
  docx = OoxmlParser::DocxParser.parse_docx(file_path)
  docx.background.color1 == OoxmlParser::Color.new(255, 255, 255)
end

# Verify the file and report the results
def verify_and_report(file)
  begin
    result = verification(file)
  rescue StandardError => e
    puts "Error: #{e.message}"
    result = false
  end

  if result
    puts 'File is correct'
    exit(0)
  else
    puts 'Something is wrong with the file'
    exit(1)
  end
end

options = parse_command_line
verify_and_report(options[:file])
