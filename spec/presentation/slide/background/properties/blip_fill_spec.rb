# frozen_string_literal: true

require 'spec_helper'

describe 'BlipFill' do
  it 'stretch.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/background/properties/blip_fill/stretch.pptx')
    expect(pptx.slides[0].background.properties.blip_fill.stretch).to be_a(OoxmlParser::Stretch)
  end
end
