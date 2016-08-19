require 'spec_helper'

describe 'RunProperties' do
  it 'custom_run_properties' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/properties/run_properties/custom_run_properties.docx')
    expect(docx.elements.first.paragraph_properties.run_properties.size.value).to eq(52)
    expect(docx.elements.first.paragraph_properties.run_properties.spacing.value).to eq(5)
    expect(docx.elements.first.paragraph_properties.run_properties.color).to eq(OoxmlParser::Color.new(255, 255, 0))
  end
end
