require 'spec_helper'

describe 'My behaviour' do
  it 'default_font_size.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/paragraphs/runs/properties/font_size/default_font_size.pptx')
    expect(pptx.slides.first.nonempty_elements.first.text_body.paragraphs[0].characters[0].properties.font_size).to eq(18)
  end
end
