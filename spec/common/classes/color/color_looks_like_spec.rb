# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Color, '#looks_like' do
  let(:no_init_color) { described_class.new }
  let(:zero_color) { described_class.new(0, 0, 0) }
  let(:init_color) { described_class.new(100, 150, 200) }

  it 'Compare none with any color' do
    expect(no_init_color).not_to be_looks_like(zero_color)
  end

  it 'Compare any with nil color' do
    expect(zero_color).not_to be_looks_like(no_init_color)
  end

  it 'Color is looks like itself' do
    expect(zero_color).to be_looks_like(zero_color)
  end

  it 'Color is looks like color with small delta' do
    color1 = init_color
    color2 = init_color.copy
    color2.red = color2.red + 1
    expect(color1).to be_looks_like(color2)
  end

  it 'Color looks like nil' do
    expect(described_class.new(nil, nil, nil)).to be_looks_like(nil)
  end

  it 'Color with style nil looks like color nil' do
    color = zero_color
    color.style = :nil
    expect(no_init_color).to be_looks_like(color)
  end
end
