# frozen_string_literal: true

require 'spec_helper'

describe 'ParagraphProperties#shade' do
  docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph' \
                                   '/properties/shade/simple_shade.docx')

  shade = docx.elements.first.paragraph_properties.shade

  it 'shade is correct class by default' do
    expect(shade).to be_a(OoxmlParser::Shade)
  end

  it 'shade has correct value' do
    expect(shade.value).to eq(:clear)
  end

  it 'shade has correct color' do
    expect(shade.color).to eq(:C9C9C9)
  end

  it 'shade has correct fill' do
    expect(shade.fill).to eq(OoxmlParser::Color.new(201, 201, 201))
  end
end
