# frozen_string_literal: true

require 'spec_helper'

describe 'hover_hyperlink' do
  it 'hover_hyperlink' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/non_visual_properties/common_properties/hover_hyperlink/hover_hyperlink.pptx')
    expect(pptx).to be_with_data
  end
end
