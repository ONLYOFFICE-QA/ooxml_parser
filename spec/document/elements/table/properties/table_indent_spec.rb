require 'spec_helper'

describe 'My behaviour' do
  it 'set_table_indent' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/properties/table_indent/set_table_indent.docx')
    expect(docx.elements[1].table_properties.table_indent).to eq(OoxmlParser::OoxmlSize.new(1440))
  end

  it 'table_indent_in_cm' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/properties/table_indent/table_indent_in_cm.docx')
    expect(docx.element_by_description[1].table_properties
               .table_indent).to eq(OoxmlParser::OoxmlSize.new(1.5, :centimeter))
  end
end
