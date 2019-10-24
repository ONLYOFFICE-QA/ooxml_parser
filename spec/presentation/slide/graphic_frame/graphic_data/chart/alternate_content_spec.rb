# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'ChartOnSlide' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/graphic_frame/graphic_data/chart/altername_content/chart_on_slide.pptx')
    expect(pptx.slides.first.elements.first.graphic_data.first.alternate_content).not_to be_nil
  end
end
