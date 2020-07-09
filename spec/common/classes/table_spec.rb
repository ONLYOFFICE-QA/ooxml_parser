# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Table do
  let(:table) { described_class.new(%w[foo bar]) }

  it '#to_s result rows data' do
    expect(table.to_s).to eq('Rows: foo,bar')
  end

  it '#inspec get same result as to_s' do
    expect(table.inspect).to eq(table.to_s)
  end
end
