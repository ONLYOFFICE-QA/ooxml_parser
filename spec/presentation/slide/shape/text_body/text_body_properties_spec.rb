require 'spec_helper'

describe 'My behaviour' do
  it 'Check VerticalAlign Top' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/properties/vertical_align_top.pptx')
    expect(pptx.slides[0].nonempty_elements.first.text_body.properties.vertical_align).to eq(:top)
  end

  it 'Check VerticalAlign bottom' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/properties/vertical_align_bottom.pptx')
    expect(pptx.slides[0].nonempty_elements.first.text_body.properties.vertical_align).to eq(:bottom)
  end
end
