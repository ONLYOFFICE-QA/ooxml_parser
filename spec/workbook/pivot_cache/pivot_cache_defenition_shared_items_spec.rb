# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::PivotCacheDefinition do
  let(:pivot_cache_parsed) { OoxmlParser::Parser.parse('spec/workbook/pivot_cache/pivot_cache.xlsx') }
  let(:shared_items) { pivot_cache_parsed.pivot_caches[0].pivot_cache_definition.cache_fields[0].shared_items }

  it 'Shared items parsing contains_semi_mixed_types' do
    expect(shared_items.contains_semi_mixed_types).to be_falsey
  end

  it 'Shared items parsing contains_string' do
    expect(shared_items.contains_string).to be_falsey
  end

  it 'Shared items parsing contains_number' do
    expect(shared_items.contains_number).to be_truthy
  end

  it 'Shared items parsing contains_integer' do
    expect(shared_items.contains_integer).to be_truthy
  end

  it 'Shared items parsing min_value' do
    expect(shared_items.min_value).to eq(1)
  end

  it 'Shared items parsing max_value' do
    expect(shared_items.max_value).to eq(2)
  end
end
