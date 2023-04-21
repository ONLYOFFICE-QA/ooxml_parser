# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Color, '#within_delta?' do
  let(:red) { described_class.new(255, 0, 0) }
  let(:green) { described_class.new(0, 255, 0) }

  it 'returns true when the colors are within the delta' do
    expect(red).to be_within_delta(green, 260)
  end

  it 'returns false when the colors are not within the delta' do
    expect(red).not_to be_within_delta(green, 10)
  end
end
