# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'numbering.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/paragraphs/properties/numbering/numbering.pptx')
    expect(pptx.slides.first.nonempty_elements.first.text_body.paragraphs[0]
               .properties.numbering.type).to eq(:alphaUcPeriod)
  end

  it 'bullet.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/paragraphs/properties/numbering/bullet.pptx')
    expect(pptx.slides.first.nonempty_elements.first.text_body.paragraphs[0]
               .properties.numbering.symbol).to eq('•')
  end

  it 'image.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/paragraphs/properties/numbering/image.pptx')
    expect(pptx.slides.first.nonempty_elements.first.text_body.paragraphs[0]
               .properties.numbering.image.file_reference.resource_id).to eq('rId2')
  end
end
