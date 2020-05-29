# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Color, '.class_methods' do
  let(:color) { described_class.new(1, 2, 3) }

  it 'Color.generate_random_color' do
    expect(described_class.generate_random_color).to be_a(described_class)
  end

  describe 'Color.array_from_const' do
    let(:color_array) { described_class.array_from_const(['none', 'RGB (1, 2, 3)']) }

    it 'array_from_const can read none color' do
      expect(color_array[0]).to eq(described_class.new)
    end

    it 'array_from_const can read specific color' do
      expect(color_array[1]).to eq(described_class.new(1, 2, 3))
    end
  end

  describe 'Color.parse_string' do
    it 'Return color if Color send' do
      expect(described_class.parse_string(color)).to eq(color)
    end

    it 'Parse color with parentheses' do
      expect(described_class.parse_string('RGB (1, 2, 3)')).to eq(color)
    end

    it 'Parse color without parentheses' do
      expect(described_class.parse_string('RGB 1, 2, 3')).to eq(color)
    end

    it 'Raise exception for incorrect color string' do
      expect { described_class.parse_string('test') }.to raise_error(/Incorrect data/)
    end
  end
end
