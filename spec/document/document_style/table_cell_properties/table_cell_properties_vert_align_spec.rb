# frozen_string_literal: true

require 'spec_helper'

describe 'CellProperties#vertical_align' do
  it 'vertical_align is a nil if undefined' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/document_style/table_cell_properties/color/cell_properties_color_nil.docx')
    expect(docx.elements[1].properties.table_style.table_cell_properties.vertical_align).to be_nil
  end

  it 'vertical_align is a symbol in other cases' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/document_style/paragraph_style_id_word.docx')
    expect(docx.elements[48].rows[0].cells[0].properties.vertical_align).to eq(:center)
  end
end
