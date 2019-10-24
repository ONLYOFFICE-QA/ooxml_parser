# frozen_string_literal: true

require 'spec_helper'

describe 'TableStyle' do
  it 'table_style_height' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/document_style/table_style/table_style_height.docx')
    expect(docx.document_style_by_name('CustomTableStyle').table_row_properties.height.value).to eq(OoxmlParser::OoxmlSize.new(1440, :twip))
  end

  it 'table_style_properties_table_properties' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/document_style/table_style/table_style_properties_table_properties.docx')
    expect(docx.document_style_by_name('CustomTableStyle').table_style_properties_list.first.table_properties).to be_a(OoxmlParser::TableProperties)
  end

  it 'table_style_properties_table_row_properties' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/document_style/table_style/table_style_properties_table_row_properties.docx')
    expect(docx.document_style_by_name('CustomTableStyle')
               .table_style_properties_list
               .first.table_row_properties).to be_a(OoxmlParser::TableRowProperties)
  end
end
