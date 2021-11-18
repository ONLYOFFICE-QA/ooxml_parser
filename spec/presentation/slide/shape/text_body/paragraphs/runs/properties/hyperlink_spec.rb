# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  files_dir = 'spec/presentation/slide/shape/text_body/'\
              'paragraphs/runs/properties/hyperlink'

  it 'HyperlinkToSlide' do
    pptx = OoxmlParser::PptxParser.parse_pptx("#{files_dir}/hyperlink_to_slide.pptx")
    expect(pptx.slides.last.elements.first.text_body.paragraphs
               .first.characters.first.properties
               .hyperlink.action).to eq(:previous_slide)
  end

  it 'HyperlinkToSpecificSlide' do
    pptx = OoxmlParser::PptxParser.parse_pptx("#{files_dir}/hyperlink_to_specific_slide.pptx")
    hyperlink = pptx.slides.last.elements.first
                    .text_body.paragraphs.first
                    .characters.first.properties.hyperlink
    expect(hyperlink.action).to eq(:slide)
    expect(hyperlink.link_to).to eq(1)
  end

  it 'LinkInTextArea' do
    pptx = OoxmlParser::PptxParser.parse_pptx("#{files_dir}/link_in_text_area.pptx")
    hyperlink = pptx.slides.last.elements.first
                    .text_body.paragraphs.first
                    .characters.first.properties.hyperlink
    expect(hyperlink.action).to eq(:external_link)
    expect(hyperlink.link_to).to eq('http://www.yandex.ru/')
  end

  it 'Hyperlink without specific id should not crash parser' do
    pptx = OoxmlParser::PptxParser.parse_pptx("#{files_dir}/hyperlink_without_id.pptx")
    expect(pptx).to be_with_data
  end
end
