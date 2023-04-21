# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Color do
  let(:no_init_color) { described_class.new }
  let(:zero_color) { described_class.new(0, 0, 0) }
  let(:init_color) { described_class.new(100, 150, 200) }

  it 'Default color to_s is none' do
    expect(no_init_color.to_s).to eq('none')
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
    it 'Default color to hex return something' do
      expect(no_init_color.to_hex).to be_nil
    end

    it 'Zero color ho hex is correct' do
      expect(zero_color.to_hex).to eq('000000')
    end

    it 'Custom color to hex is correct' do
      expect(init_color.to_hex).to eq('6496C8')
    end
  end
end
