# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Fill do
  let(:fill) do
    pattern = OoxmlParser::PatternFill.new
    pattern.foreground_color = OoxmlParser::Color.new(1, 2, 3)
    fill = described_class.new
    fill.pattern_fill = pattern
    fill
  end

  it 'Color.to_color for Fill' do
    expect(OoxmlParser::Color.to_color(fill)).to eq(OoxmlParser::Color.new(1, 2, 3))
  end
end
