require 'spec_helper'

describe 'My behaviour' do
  it 'transition_direction_nil_1' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/transition/properties/transition_direction_nil_1.pptx')
    expect(pptx.slides.first.transition.properties.direction).to be_nil
  end

  it 'transition_direction_nil_2' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/transition/properties/transition_direction_nil_2.pptx')
    expect(pptx.slides[1].transition.properties.direction).to be_nil
  end

  it 'transition_no_orientation' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/transition/properties/transition_no_orientation.pptx')
    expect(pptx.slides.first.transition.properties.orientation).to be_nil
  end
end
