# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'FontColor1' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/paragraphs/runs/properties/font_color/font_color_1.pptx')
    expect(pptx.slides.first.nonempty_elements.first.text_body.paragraphs[0].characters[0].properties.font_color.color).to eq(OoxmlParser::Color.new(242, 242, 242))
  end

  it 'FontColor2' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/paragraphs/runs/properties/font_color/font_color_2.pptx')
    expect(pptx.slides.first.nonempty_elements.first.text_body.paragraphs[0].characters[0].properties.font_color.color).to eq(OoxmlParser::Color.new(165, 165, 165))
  end

  it 'FontColor3' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/paragraphs/runs/properties/font_color/font_color_3.pptx')
    expect(pptx.slides.first.nonempty_elements.first.text_body.paragraphs[0].characters[0].properties.font_color.color).to eq(OoxmlParser::Color.new(64, 64, 64))
  end
end
