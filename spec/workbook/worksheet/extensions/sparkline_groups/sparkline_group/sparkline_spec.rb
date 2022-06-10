# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Sparkline do
  let(:xlsx) { OoxmlParser::Parser.parse('spec/workbook/worksheet/extensions/sparkline_groups/sparkline_group/sparkline_group_type.xlsx') }
  let(:sparkline) { xlsx.worksheets[0].extension_list[0].sparkline_groups[0].sparklines[0] }

  it 'Has source_reference' do
    expect(sparkline.source_reference).to eq('Отчет!C6:M6')
  end

  it 'Has destination_reference' do
    expect(sparkline.destination_reference).to eq('$N$6')
  end
end
