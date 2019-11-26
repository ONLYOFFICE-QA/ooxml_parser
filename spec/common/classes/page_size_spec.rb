# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::PageSize do
  let(:page_size) do
    described_class.new(5, 10, :portrait)
  end

  it 'PageSize#to_s is correct' do
    expect(page_size.to_s).to eq('Height: 5 Width: 10 Orientation: portrait')
  end
end
