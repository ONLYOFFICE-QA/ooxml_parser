# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'transition_direction.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/alternate_content/properties/direction/transition_direction.pptx')
    expect(pptx.slides.first.alternate_content.transition.properties.direction).to eq(:right_down)
  end

  it 'transition_direction_down.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/alternate_content/properties/direction/transition_direction_down.pptx')
    expect(pptx.slides.first.alternate_content.transition.properties.direction).to eq(:down)
  end

  it 'transition_direction_up.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/alternate_content/properties/direction/transition_direction_up.pptx')
    expect(pptx.slides.first.alternate_content.transition.properties.direction).to eq(:up)
  end

  it 'transition_direction_in.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/alternate_content/properties/direction/transition_direction_in.pptx')
    expect(pptx.slides.first.alternate_content.transition.properties.direction).to eq(:in)
  end
end
