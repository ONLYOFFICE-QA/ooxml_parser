# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'table_cell_width_size' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/table/cell/properties/table_cell_width/table_cell_width_size.docx')
    expect(docx.elements[1].rows[0].cells[0].cell_properties.table_cell_width).to eq(OoxmlParser::OoxmlSize.new(8.44, :centimeter))
  end
end
