# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  let(:pptx) { OoxmlParser::PptxParser.parse_pptx('spec/presentation/_files/empty_file.pptx') }

  it 'slide_layouts is an array' do
    expect(pptx.slide_layouts).to be_a(Array)
  end

  it 'slide_layouts element is SlideLayouts' do
    expect(pptx.slide_layouts[0]).to be_a(OoxmlParser::SlideLayoutFile)
  end

  it 'slide_master element contains common_slide_data' do
    expect(pptx.slide_layouts[0].common_slide_data).to be_a(OoxmlParser::CommonSlideData)
  end

  it 'check that excpetion is raised if slide layout file is broken' do
    expect { OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide_layouts/broken_slide_layout_file.pptx') }
      .to raise_error(OoxmlParser::NokogiriParsingException, /scene3d/)
  end
end
