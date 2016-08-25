require 'spec_helper'

describe 'Slide' do
  it 'image_horizontal_alignment_center' do
    pptx = OoxmlParser::Parser.parse('spec/presentation/slide/image_horizontal_alignment_center.pptx')
    expect(pptx.slides.first.content_horizontal_align(:image, pptx.slide_size)).to eq(:center)
  end

  it 'image_vertical_alignment_middle' do
    pptx = OoxmlParser::Parser.parse('spec/presentation/slide/image_vertical_alignment_middle.pptx')
    expect(pptx.slides.first.content_vertical_align(:image, pptx.slide_size)).to eq(:middle)
  end
end
