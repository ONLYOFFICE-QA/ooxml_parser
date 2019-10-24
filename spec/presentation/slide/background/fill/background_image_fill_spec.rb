# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'ImageFill' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/background/fill/image/image_fill.pptx')
    expect(pptx.slides[0].background.fill.image.tile).not_to be_nil
  end

  it 'image_embeded_by_link' do
    expect { OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/background/fill/image/image_no_embedded.pptx') }.not_to raise_error
  end
end
