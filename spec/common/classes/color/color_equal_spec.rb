# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Color, '#==' do
  let(:zero_color) { described_class.new(0, 0, 0) }
  let(:white_color) { described_class.new(255, 255, 255) }

  it 'Copied color is the same as original' do
    color = zero_color
    color2 = color.copy
    expect(color).to eq(color2)
  end

  it 'compare none color with whine' do
    expect(described_class.new).to eq(white_color)
  end
end
