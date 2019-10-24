# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'Align left' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/paragraphs/runs/properties/align/align_left.pptx')
    expect(pptx.slides.first.elements.last.text_body.paragraphs.first.properties.align).to eq(:left)
  end
end
