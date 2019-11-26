# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Color do
  let(:no_init_color) { OoxmlParser::Color.new }
  let(:zero_color) { OoxmlParser::Color.new(0, 0, 0) }
  let(:init_color) { OoxmlParser::Color.new(100, 150, 200) }

  it 'Color looks like nil' do
    expect(OoxmlParser::Color.new(nil, nil, nil)).to be_looks_like(nil)
  end

  it 'Color with style nil looks like color nil' do
    color = zero_color
    color.style = :nil
    expect(no_init_color).to be_looks_like(color)
  end

  it 'Default color to_s is none' do
    expect(no_init_color.to_s).to eq('none')
  end

  it 'Copied color is the same as original' do
    color = zero_color
    color2 = color.copy
    expect(color).to eq(color2)
  end

  describe 'Color#any?' do
    it 'Default color is not any?' do
      expect(no_init_color).not_to be_any
    end

    it 'Custom color is any?' do
      expect(zero_color).to be_any
    end
  end

  describe 'Color#to_hex' do
    it 'Zero color ho hex is correct' do
      expect(zero_color.to_hex).to eq('000000')
    end

    it 'Custom color to hex is correct' do
      expect(init_color.to_hex).to eq('6496C8')
    end
  end

  describe 'Color#looks_like' do
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
  end
end
