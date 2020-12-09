# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::PivotCacheDefinition do
  let(:pivot_cache_parsed) { OoxmlParser::Parser.parse('spec/workbook/pivot_cache/pivot_cache.xlsx') }
  let(:cache_fields) { pivot_cache_parsed.pivot_caches[0].pivot_cache_definition.cache_fields }

  it 'cache fields count is correct' do
    expect(cache_fields.count).to eq(1)
  end

  it 'cache fields can access cache field' do
    expect(cache_fields[0]).to be_a(OoxmlParser::CacheField)
  end
end
