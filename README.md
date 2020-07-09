# ooxml_parser

_ooxml_parser_ is a Ooxml files (docx, xlsx, pptx) parser written in Ruby.

## Installation

1. Install `magic` lib, required by `ruby-filemagic`  
   Mac OS:

   ```shell script
   brew install libmagic
   ```  

   Debian-Based Linux:

   ```shell script
   sudo apt-get install libmagic-dev
   ```

   Centos-Based Linux:

   ```shell script
   sudo yum install file-devel
   ```  

2. Install gem by command:

    ```shell script
    gem install ooxml_parser
    ```  

## Usage

Parse a docx file

```ruby
require 'ooxml_parser'
docx = OoxmlParser::Parser.parse('spec/document/document_properties/page_count.docx')
p docx.document_properties.pages # 2
```

## Configuration

### Accuracy

Accuracy of digits in fraction part
Default is 2 digits in fraction part

```ruby
OoxmlParser.configure do |config|
  config.accuracy = 3
end
