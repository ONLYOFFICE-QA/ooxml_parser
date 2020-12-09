# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::PivotCacheDefinition do
  let(:pivot_cache_parsed) { OoxmlParser::Parser.parse('spec/workbook/pivot_cache/pivot_cache.xlsx') }
  let(:cache_field) { pivot_cache_parsed.pivot_caches[0].pivot_cache_definition.cache_fields[0] }

  it 'CacheField contains name' do
    expect(cache_field.name).to eq('Column 1')
  end

  it 'CacheField contains number format id' do
    expect(cache_field.number_format_id).to eq(0)
  end

  it 'CacheField contains shared items' do
    expect(cache_field.shared_items).to be_a(OoxmlParser::SharedItems)
  end
end
