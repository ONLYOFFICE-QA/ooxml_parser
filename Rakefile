require 'rspec/core/rake_task'

desc 'Test All'
task :test do
  RSpec::Core::RakeTask.new(:spec) do |t|
    t.pattern = 'spec/*_spec.rb'
  end
  Rake::Task['spec'].execute
end
