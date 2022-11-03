# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Color, '#parse_hex_string' do
  it 'parsing hex string with auto do nothing' do
    expect(described_class.new.parse_hex_string('auto'))
      .to eq(described_class.new)
  end

  it 'parsing hex string with 6 symbols' do
    expect(described_class.new.parse_hex_string('ABCDEF').to_s)
      .to eq('RGB (171, 205, 239)')
  end

  describe 'with 8 symbols' do
    let(:with_alpha) { described_class.new.parse_hex_string('12ABCDEF') }

    it 'rgb values are correct' do
      expect(with_alpha.to_s)
        .to eq('RGB (171, 205, 239)')
    end

    it 'alpha value is correct' do
      expect(with_alpha.alpha_channel)
        .to eq(18)
    end
  end
end
