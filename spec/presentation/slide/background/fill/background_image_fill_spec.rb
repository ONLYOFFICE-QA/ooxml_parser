require 'spec_helper'

describe 'My behaviour' do
  it 'ImageFill' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/background/fill/image/image_fill.pptx')
    expect(pptx.slides[0].background.fill.image.tile).not_to be_nil
  end

  describe 'stretching' do
    it 'fill_rectangle_strectching' do
      pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/background/fill/image/stretching/fill_rectangle_strectching.pptx')
      expect(pptx.slides.first.background.fill.image.stretch.fill_rectangle).to be_a(OoxmlParser::FillRectangle)
    end
  end

  it 'image_no_embedded' do
    expect { OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/background/fill/image/image_no_embedded.pptx') }.to raise_error(LoadError)
  end
end
