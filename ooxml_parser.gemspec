# frozen_string_literal: true

gemspec = instance_eval(File.read(File.expand_path('ooxml_parser-mingw32.gemspec', __dir__)))

gemspec.platform = Gem::Platform::RUBY
gemspec.required_ruby_version = '>= 2.5'
gemspec.add_dependency 'ruby-filemagic', '~> 0'
gemspec
