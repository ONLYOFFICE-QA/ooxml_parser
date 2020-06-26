# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::PivotCacheDefinition do
  let(:pivot_cache_parsed) { OoxmlParser::Parser.parse('spec/workbook/pivot_cache/pivot_cache.xlsx') }
  let(:cache_source) { pivot_cache_parsed.pivot_caches[0].pivot_cache_definition.cache_source }
  let(:worksheet_source) { cache_source.worksheet_source }

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

  it 'worksheet_source class is correct' do
    expect(worksheet_source).to be_a(OoxmlParser::WorksheetSource)
  end

  it 'worksheet_source has correct ref' do
    expect(worksheet_source.ref).to eq('A1:A3')
  end

  it 'worksheet_source has correct sheet' do
    expect(worksheet_source.sheet).to eq('Лист1')
  end
end
