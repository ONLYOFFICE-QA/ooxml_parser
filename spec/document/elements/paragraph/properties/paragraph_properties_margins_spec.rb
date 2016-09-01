require 'spec_helper'

describe 'ParagraphProperties Margins' do
  it 'paragraph_properties_margin_left' do
    pptx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/properties/margins/paragraph_properties_margin_left.pptx')
    expect(pptx.slides[0].nonempty_elements[0].text_body.paragraphs[0].properties.margin_left).to eq(OoxmlParser::OoxmlSize.new(2.02, :centimeter))
  end

  it 'paragraph_properties_margin_right' do
    pptx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/properties/margins/paragraph_properties_margin_right.pptx')
    expect(pptx.slides[0].nonempty_elements[0].text_body.paragraphs[0].properties.margin_right).to eq(OoxmlParser::OoxmlSize.new(3.03, :centimeter))
  end
end
