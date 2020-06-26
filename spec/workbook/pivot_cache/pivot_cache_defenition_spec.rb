# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::PivotCacheDefinition do
  let(:pivot_cache_parsed) { OoxmlParser::Parser.parse('spec/workbook/pivot_cache/pivot_cache.xlsx') }
  let(:cache_source) { pivot_cache_parsed.pivot_caches[0].pivot_cache_definition.cache_source }

  it 'Pivot cache definition can be found for pivot cache' do
    expect(pivot_cache_parsed.pivot_caches[0].pivot_cache_definition)
      .to be_a(described_class)
  end

  it 'Pivot cache definition contains cache_source' do
    expect(pivot_cache_parsed.pivot_caches[0].pivot_cache_definition.cache_source)
      .to be_a(OoxmlParser::CacheSource)
  end

  it 'cache_source type is worksheet' do
    expect(cache_source.type).to eq('worksheet')
  end
end
