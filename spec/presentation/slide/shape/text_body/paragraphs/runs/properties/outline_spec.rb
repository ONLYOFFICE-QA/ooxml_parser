# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'Align left' do
    pptx = OoxmlParser::Parser.parse('spec/presentation/slide/shape/text_body/paragraphs/runs/properties/outline/default_outline_width.pptx')
    expect(pptx.slides.first.elements.last.text_body.paragraphs.first
               .runs.first.properties.outline.width).to eq(OoxmlParser::OoxmlSize.new(0, :centimeter))
  end
end
