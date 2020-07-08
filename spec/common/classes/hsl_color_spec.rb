# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::HSLColor do
  let(:hsl_color) { described_class.new }

  describe '#calculate_lum_value' do
    it 'calculate_lum_value for nil tint is lum' do
      expect(hsl_color.calculate_lum_value(nil, 50)).to eq(50)
    end
  end
end
