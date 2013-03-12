#!/opt/nite/bin/perl

use strict;
use Data::Dumper;

my %cpus;
my @times;

#my $counter = 21;

my $file = @ARGV[0];

open FILE, "$file" or die "error opening $file: $!";

while (<FILE>) {
	chomp;
	if ( $_ =~ /Usage: CPU/ ) {

		my @fields = split /\s+/, $_;

		my $cpu = $fields[2];
		if ( $_ =~ /Usage: CPU_Total/ ) {
			push( @times, $fields[0] );
		}
		my $perc = $fields[5];

		if ( ref $cpus{$cpu} ne 'ARRAY' ) {
			$cpus{$cpu} = [];
		}
		push @{ $cpus{$cpu} }, "$perc";

		#$counter --;
	}
}
close FILE;

#last if $counter <= 1;
print "Time,";
foreach my $c ( sort keys %cpus ) {
	print "$c,"
}
print "Subtotal\n";

for ( my $t = 0 ; $t < @times ; $t++ ) {
	my $total = 0;
	print $times[$t];
	foreach my $c ( sort keys %cpus ) {
		print ",", $cpus{$c}[$t];
		next if $c =~ /CPU_Total/;
		$total += $cpus{$c}[$t];
	}
	print ",$total%\n";
}
