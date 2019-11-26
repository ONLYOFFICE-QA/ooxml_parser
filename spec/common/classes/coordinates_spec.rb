# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Coordinates do
  let(:first_coord) { OoxmlParser::Coordinates.new(1, 'A') }
  let(:second_coord) { OoxmlParser::Coordinates.new(2, 'B') }

  it 'Coordinates#row_greater_that_other?' do
    expect(first_coord).not_to be_row_greater_that_other(second_coord)
  end

  it 'Coordinates#column_greater_that_other?' do
    expect(second_coord).to be_column_greater_that_other(first_coord)
  end

  it 'Coordinates#to_s' do
    expect(first_coord.to_s).to eq('A1 ')
  end

  it 'Coordinates#nil?' do
    expect(first_coord).not_to be_nil
  end
end
