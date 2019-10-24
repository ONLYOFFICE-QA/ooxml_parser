# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::CommonNonVisualProperties do
  it 'CommonNonVisualProperties#title' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/graphic_frame/non_visual_properties/common_properties/property_title.pptx')
    expect(pptx.slides.last.elements.last.non_visual_properties.common_properties.title).to eq('Simple Test Text')
  end

  it 'CommonNonVisualProperties#description' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/graphic_frame/non_visual_properties/common_properties/property_title.pptx')
    expect(pptx.slides.last.elements.last.non_visual_properties.common_properties.description).to eq('SimpleTestText')
  end
end
