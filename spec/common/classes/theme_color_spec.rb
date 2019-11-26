# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::ThemeColor do
  let(:system_color) do
    described_class.new(type: :system,
                        color: OoxmlParser::Color.new(100,
                                                      150,
                                                      200))
  end

  describe 'ThemeColor equality' do
    it 'ThemeColor is equal to same theme color' do
      expect(system_color).to eq(system_color)
    end

    it 'ThemeColor is not equal to same theme color' do
      other_color = described_class.new(type: :rgb)
      expect(system_color).not_to eq(other_color)
    end

    it 'ThemeColor is equal color of itself' do
      expect(system_color).to eq(system_color.color)
    end
  end
end
