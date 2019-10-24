# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'table_cell_spacing' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/table/properties/table_cell_spacing/table_cell_spacing.docx')
    expect(docx.elements[1].properties.table_cell_spacing).to eq(OoxmlParser::OoxmlSize.new(1.27, :centimeter))
  end
end
