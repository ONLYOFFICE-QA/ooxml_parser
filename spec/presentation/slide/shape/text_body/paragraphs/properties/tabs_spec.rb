require 'spec_helper'

describe 'My behaviour' do
  it 'numbering.pptx' do
    pptx = OoxmlParser::Parser.parse('spec/presentation/slide/shape/text_body/paragraphs/properties/tabs/tabs.pptx')
    expect(pptx.slides[0].elements.last.text_body.paragraphs[0].properties
                     .tabs[0].position).to eq(OoxmlParser::OoxmlSize.new(1.25, :centimeter))
    expect(pptx.slides[0].elements.last.text_body.paragraphs[0].properties
                     .tabs[0].align).to eq(:left)
  end
end
