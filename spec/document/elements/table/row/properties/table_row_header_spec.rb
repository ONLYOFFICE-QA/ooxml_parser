# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'table_row_header_enabled' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/row/properties/table_header/table_row_header_enabled.docx')
    expect(docx.elements[1].rows.first.table_row_properties.table_header).to be_truthy
  end

  it 'table_row_header_disabled' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/table/row/properties/table_header/table_row_header_disabled.docx')
    expect(docx.elements[1].rows.first.table_row_properties.table_header).to be_falsey
  end
end
