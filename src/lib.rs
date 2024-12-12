mod repack;

use repack::example_usage;
pub fn add(left: u64, right: u64) -> u64 {
    left + right
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn it_works() {
        example_usage();
        let result = add(2, 2);
        assert_eq!(result, 4);
    }
}
