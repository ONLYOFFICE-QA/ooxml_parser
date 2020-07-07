# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::DocxColorScheme do
  let(:color_scheme) { described_class.new }

  it '#to_s output data for default scheme' do
    expect(color_scheme.to_s).to eq('Color: none, type: unknown')
  end
end
