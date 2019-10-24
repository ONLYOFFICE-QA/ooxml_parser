# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'custom_table_width' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/table/properties/table_width/custom_table_width.docx')
    expect(docx.elements[1].table_properties.table_width).to eq(OoxmlParser::OoxmlSize.new(12.7, :centimeter))
  end

  it 'table_width_in_percent' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/table/properties/table_width/table_width_in_percent.docx')
    expect(docx.elements[1].table_properties.table_width).to eq(OoxmlParser::OoxmlSize.new(100, :percent))
  end
end
