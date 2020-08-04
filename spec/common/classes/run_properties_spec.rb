# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::RunProperties do
  it 'RunProperties can be initialized with specific font name' do
    font_name = 'Arial'
    expect(described_class.new(font_name: font_name).font_name).to eq(font_name)
  end
end
