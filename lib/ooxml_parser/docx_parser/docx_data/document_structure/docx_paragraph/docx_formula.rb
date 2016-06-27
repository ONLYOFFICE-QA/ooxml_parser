# Class for containg DocxFormulas
require_relative 'docx_formula/accent'
require_relative 'docx_formula/argument_properties'
require_relative 'docx_formula/bar'
require_relative 'docx_formula/box'
require_relative 'docx_formula/delimeter'
require_relative 'docx_formula/fraction'
require_relative 'docx_formula/function'
require_relative 'docx_formula/group_char'
require_relative 'docx_formula/limit'
require_relative 'docx_formula/math_run'
require_relative 'docx_formula/matrix'
require_relative 'docx_formula/nary'
require_relative 'docx_formula/operator'
require_relative 'docx_formula/index'
require_relative 'docx_formula/radical'
module OoxmlParser
  class DocxFormula
    attr_accessor :formula_run
    # @return [ArgumentProperties] properties of arguments
    attr_accessor :argument_properties

    def initialize(elements = [])
      @formula_run = elements
    end

    def self.parse(element)
      formula = DocxFormula.new
      element.xpath('*').each do |sub_element|
        if sub_element.name == 'r'
          formula.formula_run << MathRun.parse(sub_element)
        elsif sub_element.name == 'box'
          box = Box.new
          box.element = MathRun.parse(sub_element)
          formula.formula_run << box
        elsif sub_element.name == 'borderBox'
          box = Box.new
          box.borders = true
          box.element = MathRun.parse(sub_element)
          formula.formula_run << box
        elsif sub_element.name == 'func'
          function = Function.new
          sub_element.xpath('m:fName').each do |f_name|
            function.name = DocxFormula.parse(f_name)
          end
          function.argument = DocxFormula.parse(sub_element)
          formula.formula_run << function
        elsif sub_element.name == 'rad'
          radical = Radical.new
          radical.value = DocxFormula.parse(sub_element)
          sub_element.xpath('m:deg').each do |deg|
            radical.degree = DocxFormula.parse(deg)
            radical.degree = 2 if radical.degree.nil?
          end
          formula.formula_run << radical
        elsif sub_element.name == 'e'
          formula.formula_run << DocxFormula.parse(sub_element)
        elsif sub_element.name == 'eqArr'
          sub_element.xpath('m:e').each do |e|
            formula.formula_run << DocxFormula.parse(e)
          end
        elsif sub_element.name == 'd'
          delimeter = Delimiter.new
          sub_element.xpath('m:dPr').each do |d_pr|
            d_pr.xpath('m:begChr').each do |beg_chr|
              delimeter.begin_character = beg_chr.attribute('val').value
            end
            d_pr.xpath('m:endChr').each do |end_chr|
              delimeter.end_character = end_chr.attribute('val').value
            end
          end
          delimeter.value = DocxFormula.parse(sub_element)
          formula.formula_run << delimeter
        elsif sub_element.name == 'nary'
          formula.formula_run << Nary.parse(sub_element)
        elsif sub_element.name == 'sSubSup'
          index = Index.new
          index.value = DocxFormula.parse(sub_element)
          sub_element.xpath('m:sup').each do |sup|
            index.top_index = DocxFormula.parse(sup)
          end
          sub_element.xpath('m:sub').each do |sub|
            index.top_index = DocxFormula.parse(sub)
          end
          formula.formula_run << index
        elsif sub_element.name == 'sSup'
          index = Index.new
          index.value = DocxFormula.parse(sub_element)
          sub_element.xpath('m:sup').each do |sup|
            index.top_index = DocxFormula.parse(sup)
          end
          formula.formula_run << index
        elsif sub_element.name == 'sSub'
          index = Index.new
          index.value = DocxFormula.parse(sub_element)
          sub_element.xpath('m:sub').each do |sub|
            index.top_index = DocxFormula.parse(sub)
          end
          formula.formula_run << index
        elsif sub_element.name == 'f'
          fraction = Fraction.new
          sub_element.xpath('m:num').each do |num|
            fraction.numerator = DocxFormula.parse(num)
          end
          sub_element.xpath('m:den').each do |den|
            fraction.denominator = DocxFormula.parse(den)
            formula.formula_run << fraction
          end
        elsif sub_element.name == 'm'
          matrix = Matrix.new
          columns_count = 1
          j = 0
          sub_element.xpath('m:mPr').each do |m_pr|
            m_pr.xpath('m:mcs').each do |mcs|
              mcs.xpath('m:mc').each do |mc|
                mc.xpath('m:mcPr').each do |mc_pr|
                  mc_pr.xpath('m:count').each do |count|
                    columns_count = count.attribute('val').value
                  end
                end
              end
            end
          end
          sub_element.xpath('m:mr').each do |mr|
            i = 0
            matrix.rows << MatrixRow.new(columns_count.to_i)
            mr.xpath('m:e').each do |e|
              matrix.rows[j].columns[i] = DocxFormula.parse(e)
              i += 1
            end
            j += 1
          end
          formula.formula_run << matrix
        elsif sub_element.name == 'bar'
          bar = Bar.new
          sub_element.xpath('m:barPr').each do |bar_pr|
            bar_pr.xpath('m:pos').each do |pos|
              bar.position = pos.attribute('val').value
            end
          end
          bar.element = DocxFormula.parse(sub_element)
          formula.formula_run << bar
        elsif sub_element.name == 'acc'
          accent = Accent.new
          sub_element.xpath('m:accPr').each do |acc_pr|
            acc_pr.xpath('m:chr').each do |chr|
              accent.symbol = chr.attribute('val').value
            end
          end
          accent.element = DocxFormula.parse(sub_element)
          formula.formula_run << accent
        elsif sub_element.name == 'groupChr'
          group_char = GroupChar.new
          sub_element.xpath('m:groupChrPr').each do |group_chr_pr|
            group_chr_pr.xpath('m:chr').each do |chr|
              group_char.symbol = chr.attribute('val').value
            end
            group_chr_pr.xpath('m:pos').each do |pos|
              group_char.position = pos.attribute('val').value
            end
            group_chr_pr.xpath('vertJc').each do |vert_jc|
              group_char.vertical_align = vert_jc.attribute('val').value
            end
          end
          group_char.element = DocxFormula.parse(sub_element)
          formula.formula_run << group_char
        elsif sub_element.name == 'limUpp'
          limit = Limit.new
          limit.type = 'Upper'
          limit.element = DocxFormula.parse(sub_element)
          sub_element.xpath('m:lim').each do |lim|
            limit.limit = DocxFormula.parse(lim)
          end
          formula.formula_run << limit
        elsif sub_element.name == 'limLow'
          limit = Limit.new
          limit.type = 'Lower'
          limit.element = DocxFormula.parse(sub_element)
          sub_element.xpath('m:lim').each do |lim|
            limit.limit = DocxFormula.parse(lim)
          end
          formula.formula_run << limit
        elsif sub_element.name == 'argPr'
          formula.argument_properties = ArgumentProperties.parse(sub_element)
        end
      end
      return nil if formula.formula_run.empty?
      formula
    end
  end
end
