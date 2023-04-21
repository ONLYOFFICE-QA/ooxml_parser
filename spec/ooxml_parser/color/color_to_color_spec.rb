# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Color, '.to_color' do
  it 'SchemeColor to_color' do
    object = OoxmlParser::SchemeColor.new(nil,
                                          'scheme')
    expect(described_class.to_color(object)).to eq('scheme')
  end

  it 'PresenationFill with converted color to_color' do
    object = OoxmlParser::PresentationFill.new
    object.color = OoxmlParser::SchemeColor.new(nil,
                                                'scheme')
    expect(described_class.to_color(object)).to eq('scheme')
  end

  it 'DocxColorScheme to_color' do
    object = OoxmlParser::DocxColorScheme.new
    expect(described_class.to_color(object)).to eq(described_class.new)
  end

  it 'String to_color' do
    object = 'none'
    expect(described_class.to_color(object)).to eq(described_class.new)
  end

  it 'Symbol to_color' do
    object = :none
    expect(described_class.to_color(object)).to eq(described_class.new)
  end

  it 'DocxColor to_color' do
    object = OoxmlParser::DocxColor.new
    object.value = described_class.new
    expect(described_class.to_color(object)).to eq(described_class.new)
  end

  it 'Other class to_color' do
    object = Object.new
    expect(described_class.to_color(object)).to eq(object)
  end
end
