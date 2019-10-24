# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'HyperlinkToSlide' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/paragraphs/runs/properties/hyperlink/hyperlink_to_slide.pptx')
    expect(pptx.slides.last.elements.first.text_body.paragraphs.first.characters.first.properties.hyperlink.action).to eq(:previous_slide)
  end

  it 'HyperlinkToSpecificSlide' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/paragraphs/runs/properties/hyperlink/hyperlink_to_specific_slide.pptx')
    expect(pptx.slides.last.elements.first.text_body.paragraphs.first.characters.first.properties.hyperlink.action).to eq(:slide)
    expect(pptx.slides.last.elements.first.text_body.paragraphs.first.characters.first.properties.hyperlink.link_to).to eq(1)
  end

  it 'LinkInTextArea' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/paragraphs/runs/properties/hyperlink/link_in_text_area.pptx')
    expect(pptx.slides.last.elements.first.text_body.paragraphs.first.characters.first.properties.hyperlink.action).to eq(:external_link)
    expect(pptx.slides.last.elements.first.text_body.paragraphs.first.characters.first.properties.hyperlink.link_to).to eq('http://www.yandex.ru/')
  end
end
