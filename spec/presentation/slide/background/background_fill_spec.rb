require 'spec_helper'

describe 'My behaviour' do
  it 'slide_pattern.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/background/fill/slide_pattern.pptx')
    expect(pptx.slides[0].background.fill.pattern
               .background_color.value).to eq(OoxmlParser::Color.new(231, 230, 230))
  end

  it 'slide_gradient.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/background/fill/slide_gradient.pptx')
    gradient_stops = pptx.slides[0].background.fill.color.gradient_stops
    expect(gradient_stops[0].color.value).to eq(OoxmlParser::Color.new(237, 125, 49))
    expect(gradient_stops[0].position).to eq(10)
    expect(gradient_stops[1].position).to eq(70)
  end
end
