# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::PivotCache do
  pivot_cache_parsed = OoxmlParser::Parser.parse('spec/workbook/pivot_cache/pivot_cache.xlsx')

  it 'pivot cache is empty array for simplest document' do
    xlsx = OoxmlParser::Parser.parse('spec/workbook/chartsheets/file_with_chartsheets.xlsx')
    expect(xlsx.pivot_caches).to eq([])
  end

  it 'pivot cache contains cacheId' do
    expect(pivot_cache_parsed.pivot_caches[0].cache_id).to eq(1)
  end

  it 'pivot cache contains id' do
    expect(pivot_cache_parsed.pivot_caches[0].id).to eq('rId3')
  end
end
