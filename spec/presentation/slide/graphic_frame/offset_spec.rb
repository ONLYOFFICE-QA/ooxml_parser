# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'ChartTransformOffset' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/graphic_frame/offset/transform_offset.pptx')
    expect(pptx.slides.last.elements.last.transform.offset.y).to be_zero
  end

  it 'Check image size' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/graphic_frame/offset/image_size.pptx')
    transform = pptx.slides.first.elements.last.properties.transform
    expect(transform.extents.x).to eq(OoxmlParser::OoxmlSize.new(13, :centimeter))
    expect(transform.extents.y).to eq(OoxmlParser::OoxmlSize.new(5, :centimeter))
  end

  it 'Check image shift' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/graphic_frame/offset/image_shift.pptx')
    transform = pptx.slides.first.elements.last.properties.transform
    expect(transform.offset.x).to eq(pptx.slide_size.width)
    expect(transform.offset.y).to eq(pptx.slide_size.height)
  end
end
