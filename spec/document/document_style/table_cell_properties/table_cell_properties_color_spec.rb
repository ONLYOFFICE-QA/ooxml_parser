# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'table_cell_properties_border' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/document_style/table_cell_properties/color/cell_properties_color_nil.docx')
    expect(docx.elements[1].properties.table_style.table_cell_properties.color).to be_nil
  end
end
