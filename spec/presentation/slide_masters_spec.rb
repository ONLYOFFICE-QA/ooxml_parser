# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/_files/empty_file.pptx')

  it 'slide_masters is an array' do
    expect(pptx.slide_masters).to be_a(Array)
  end

  it 'slide_masters element is SlideMaster' do
    expect(pptx.slide_masters[0]).to be_a(OoxmlParser::SlideMasterFile)
  end

  it 'slide_master element contains common_slide_data' do
    expect(pptx.slide_masters[0].common_slide_data).to be_a(OoxmlParser::CommonSlideData)
  end
end
