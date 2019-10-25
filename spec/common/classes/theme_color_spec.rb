# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::ThemeColor do
  let(:system_color) { OoxmlParser::ThemeColor.new(:system) }
  describe 'ThemeColor equality' do
    it 'ThemeColor is equal to same theme color' do
      expect(system_color).to eq(system_color)
    end

    it 'ThemeColor is equal to same theme color' do
      other_color = OoxmlParser::ThemeColor.new(:rgb)
      expect(system_color).not_to eq(other_color)
    end
  end
end
