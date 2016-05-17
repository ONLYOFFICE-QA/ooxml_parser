# ooxml_parser

_ooxml_parser_ is a Ooxml files (docx, xlsx, pptx) parser written in Ruby.

## Installation

    $ gem install ooxml_parser
    
## Usage

Parse a docx file

    require 'ooxml_parser'
    docx = OoxmlParser::DocxParser.parse_docx('spec/docx_examples/document_properties/page_count.docx')
    p docx.document_properties.pages # 2
    
    
## Configuration

Set units of mesaurements:
Possible units are `:points`, `:centimeters`, `:dxa`, `:inches`, `:emu`
Default is `:centimeters`

    OoxmlParser.configure do |config|
      config.units = :points
    end
