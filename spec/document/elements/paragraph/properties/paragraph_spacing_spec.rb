require 'spec_helper'

describe 'My behaviour' do
  it 'paragraph_spacing_value_size' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/properties/spacing/paragraph_spacing_value_size.docx')
    expect(docx.elements.first.paragraph_properties.spacing.after).to eq(OoxmlParser::OoxmlSize.new(2.54, :centimeter))
  end

  it 'paragraph_spacing_line_value' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/properties/spacing/paragraph_spacing_line_value.docx')
    expect(docx.elements.first.paragraph_properties.spacing.line).to eq(OoxmlParser::OoxmlSize.new(3, :centimeter))
  end
end
