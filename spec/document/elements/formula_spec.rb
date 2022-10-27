# frozen_string_literal: true

require 'spec_helper'

describe 'formula' do
  describe 'argument_properties' do
    it 'argument_properties_size' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/formula/argument_properties/argument_properties_size.docx')
      expect(docx.elements.first.nonempty_runs.first.formula_run[1].top_index.argument_properties.argument_size.value).to eq(-1)
    end
  end

  describe 'run' do
    describe 'run_properties' do
      it 'formula_manual_break' do
        docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/formula/run/run_properties/formula_manual_break.docx')
        expect(docx.elements.first.nonempty_runs.first.formula_run.first.properties.break).to be_truthy
        expect(docx.elements.first.nonempty_runs.first.formula_run.first.text).to eq('±')
      end

      it 'formula_with_bottom_index.docx' do
        docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/formula/run/run_properties/formula_with_bottom_index.docx')
        expect(docx.elements.first.nonempty_runs.first.formula_run.first.bottom_index).to be_a(OoxmlParser::DocxFormula)
      end
    end
  end

  describe 'type' do
    it 'nary_formula_type' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/formula/types/nary_formula_type.docx')
      expect(docx.elements.first.nonempty_runs.first.formula_run[1].properties.limit_location.value).to eq(:under_over)
    end

    it 'formula_complex.docx' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/formula/types/formula_complex.docx')
      expect(docx.elements.first.nonempty_runs.first).to be_a(OoxmlParser::DocxFormula)
    end

    it 'formula_symbol.docx' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/formula/types/formula_symbol.docx')
      expect(docx.elements.first.nonempty_runs.first).to be_a(OoxmlParser::DocxFormula)
    end

    it 'fraction.docx' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/formula/types/fraction.docx')
      expect(docx.elements.first.nonempty_runs.first).to be_a(OoxmlParser::DocxFormula)
    end

    it 'formula_integral.docx' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/formula/types/formula_integral.docx')
      expect(docx.elements.first.nonempty_runs.first).to be_a(OoxmlParser::DocxFormula)
    end

    it 'formula_brackets.docx' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/formula/types/formula_brackets.docx')
      expect(docx.elements.first.nonempty_runs.first).to be_a(OoxmlParser::DocxFormula)
    end

    it 'formula_function.docx' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/formula/types/formula_function.docx')
      expect(docx.elements.first.nonempty_runs.first).to be_a(OoxmlParser::DocxFormula)
    end

    it 'formula_accent.docx' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/formula/types/formula_accent.docx')
      expect(docx.elements.first.nonempty_runs.first).to be_a(OoxmlParser::DocxFormula)
    end

    it 'formula_group_char.docx' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/formula/types/formula_group_char.docx')
      expect(docx.elements.first.nonempty_runs.first).to be_a(OoxmlParser::DocxFormula)
    end

    it 'formula_limit.docx' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/formula/types/formula_limit.docx')
      expect(docx.elements.first.nonempty_runs.first).to be_a(OoxmlParser::DocxFormula)
    end

    it 'matrix.docx' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/formula/types/matrix.docx')
      expect(docx.elements.first.nonempty_runs.first).to be_a(OoxmlParser::DocxFormula)
    end
  end
end
