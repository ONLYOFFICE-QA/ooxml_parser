# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'highlight_click' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/non_visual_properties/common_properties/on_click_hyperlink/highlight_click.pptx')
    expect(pptx.slides[17].elements[3].non_visual_properties.common_properties.on_click_hyperlink.highlight_click).to be_truthy
  end
end
