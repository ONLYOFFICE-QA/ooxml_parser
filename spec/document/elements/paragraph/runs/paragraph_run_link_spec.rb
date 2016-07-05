require 'spec_helper'

describe OoxmlParser::Hyperlink do
  it 'Check Hyperlink' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/link/hyperlink.docx')
    expect(docx.elements.first.nonempty_runs.first.link.link).to eq('http://www.yandex.ru/')
  end

  it 'Check Hyperlink with tooltip' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/link/hyperlink_tooltip.docx')
    expect(docx.notes.first.elements.first.nonempty_runs.first.link.link).to eq('http://www.yandex.ru/')
    expect(docx.notes.first.elements.first.nonempty_runs.first.link.tooltip).to eq('go to www.yandex.ru')
  end

  it 'LinkXpathError' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/link/link_xpath_error.docx')
    expect(docx.elements.first.nonempty_runs.first.text).to eq('yandex')
    expect(docx.elements.first.nonempty_runs.first.link.link).to eq('http://www.yandex.ru/')
    expect(docx.elements.first.nonempty_runs.first.link.tooltip).to eq('go to www.yandex.ru')
  end

  it 'field_inside_hyperlink' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/link/field_inside_hyperlink.docx')
    expect(docx.elements.first.nonempty_runs.first.link.url).to eq('https://www.yandex.ru/')
  end

  it 'run_with_link' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/link/run_with_link.docx')
    expect(docx.elements.first.character_style_array[0].link.link).to eq('http://www.yandex.ru')
    expect(docx.elements.first.character_style_array[0].link.tooltip).to eq('go to www.yandex.ru')
  end
end
